"""Fill semantics applied to an OOXML package.

This is the engine-independent part of ``engine="fillxl"``: it decides *what*
changes, using the same shared logic as the other engines — ``_plan`` (row
arithmetic), ``_formulas`` (calculated columns and range resizing),
``_validation`` (column matching), ``_normalize`` (value rules),
``_charts`` (reference rewriting) — and applies it with the surgical editors in
:mod:`fillxl._ooxml`.

The engine never inserts worksheet rows.  Growth lands in cells verified empty,
exactly like the openpyxl engine, so nothing below the table ever moves.
"""

from __future__ import annotations

import datetime as dt
import logging
from typing import TYPE_CHECKING, Any

from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH
from openpyxl.worksheet.cell_range import CellRange

from fillxl._adapters import TabularAdapter
from fillxl._cells import header_write_error, name_destinations, named_cell_target
from fillxl._formats import needs_temporal_format, temporal_format
from fillxl._formulas import (
    ColumnFormulaEvidence,
    classify_body,
    extend_table_ranges,
    may_track_table,
    resolve_formula,
    shift_relative_rows,
    sole_reference,
    warn_about_ignored_formula_data,
)
from fillxl._normalize import (
    check_cell_value,
    check_named_cell_scalar,
    check_nonfinite_value,
    prepare_cell_value,
    prepare_table_cell,
)
from fillxl._ooxml import xmltext
from fillxl._ooxml.cells import build_cell
from fillxl._ooxml.pivots import pivot_worksheet_sources
from fillxl._ooxml.styles import StylesPart
from fillxl._ooxml.tables import TableColumnInfo, TablePart, find_table, tables_on_sheet
from fillxl._ooxml.worksheet import (
    SheetPart,
    SheetStore,
    array_range,
    split_reference,
)
from fillxl._plan import BlockedRegion, blocked_area, plan_table_fill
from fillxl._references import RangeSource, destination_references, warn_about_untracked_ranges
from fillxl._validation import validate_columns

if TYPE_CHECKING:
    from collections.abc import Sequence

    from fillxl._charts import TableChangeSpecs
    from fillxl._config import FillXLConfig
    from fillxl._ooxml.package import OoxmlPackage
    from fillxl._ooxml.workbook import SheetEntry, WorkbookPart
    from fillxl._plan import TableGeometry, TablePlan

logger = logging.getLogger(__name__)

_DATA_VALIDATION = xmltext.element_re("dataValidation")


# ── named cells ─────────────────────────────────────────────────────────────


def fill_named_cell(
    package: OoxmlPackage,
    workbook: WorkbookPart,
    name: str,
    value: Any,  # noqa: ANN401
    *,
    escape_formulas: bool,
    sheets: SheetStore,
    styles: StylesPart,
) -> None:
    """Write a scalar *value* to the cell the defined name *name* points at.

    Args:
        package: The open package.
        workbook: The workbook part, used to resolve the name.
        name: Defined name to write.
        value: Scalar value.
        escape_formulas: Whether a string starting with ``=`` is text.
        sheets: The session's sheet editors, so a sheet is parsed once.
        styles: The session's style sheet.

    Raises:
        KeyError: If the name (or the sheet it points at) does not exist.
        ValueError: If the name points at a table header row.

    """
    entry = workbook.find_defined_name(name)
    if entry is None:
        msg = f"Named cell '{name}' not found in workbook"
        raise KeyError(msg)

    check_named_cell_scalar(name, value)
    check_nonfinite_value(value, name=name)
    normalised = prepare_cell_value(value)
    check_cell_value(normalised, name=name, date1904=workbook.date1904)

    target = named_cell_target(name, name_destinations(entry.text))
    if target is None:
        return
    sheet_name, address = target
    sheet_entry = workbook.sheet(sheet_name)
    if sheet_entry is None:
        msg = f"Named cell '{name}' points at sheet '{sheet_name}', which does not exist"
        raise KeyError(msg)
    if "/worksheets/" not in sheet_entry.part:
        logger.warning("Worksheet '%s' is not a regular worksheet, skipping", sheet_name)
        return

    sheet = sheets.get(sheet_entry.part)
    row, column = split_reference(address.replace("$", ""))
    _refuse_header_write(package, sheet_entry, row, column, sheet_name, address)
    _refuse_spill_write(sheet, row, column)

    sheet.write_cell(
        row,
        column,
        build_cell(
            sheet.reference(row, column),
            normalised,
            style=_date_style(styles, sheet.style(row, column), normalised),
            escape_formulas=escape_formulas,
            epoch=_epoch(workbook),
        ),
    )
    logger.debug("Set named cell '%s' (%s!%s) = %s", name, sheet_name, address, normalised)


def _epoch(workbook: WorkbookPart) -> dt.datetime:
    """The day zero *workbook* counts dates from."""
    return MAC_EPOCH if workbook.date1904 else WINDOWS_EPOCH


def _refuse_header_write(
    package: OoxmlPackage,
    sheet_entry: SheetEntry,
    row: int,
    column: int,
    sheet_name: str,
    address: str,
) -> None:
    """Raise when the named cell is a table header (renaming it breaks the table)."""
    for table in tables_on_sheet(package, sheet_entry):
        geometry = table.geometry()
        if geometry.header_rows == 0:
            continue
        if geometry.top_row == row and geometry.first_col <= column <= geometry.last_col:
            raise header_write_error(sheet_name, address, table.display_name)


def _date_style(styles: StylesPart, style: str | None, value: Any) -> str | None:  # noqa: ANN401
    """The style to use so a date value is shown as a date.

    Mirrors the openpyxl engine: a date in a cell whose number format is not a
    date format gets the default date format; anything else keeps its style.
    """
    format_code = temporal_format(value)
    if format_code is None:
        return style
    if not needs_temporal_format(styles.number_format(style)):
        return style
    return styles.date_style(style, format_code) or style


# ── tables ──────────────────────────────────────────────────────────────────


def fill_table(
    package: OoxmlPackage,
    workbook: WorkbookPart,
    name: str,
    data: Any,  # noqa: ANN401
    config: FillXLConfig,
    *,
    sheets: SheetStore,
    styles: StylesPart,
) -> TableChangeSpecs:
    """Fill the table called *name* with *data*.

    Args:
        package: The open package.
        workbook: The workbook part.
        name: Table name.
        data: Tabular data (DataFrame, list[dict], …).
        config: Fill configuration.
        sheets: The session's sheet editors, so a sheet is parsed once.
        styles: The session's style sheet.

    Returns:
        The old and new table dimensions, for chart updates.

    Raises:
        KeyError: If the table does not exist.
        ValueError: If the table cannot expand, or the columns do not match.

    """
    adapter = TabularAdapter(data)
    sheet_entries = workbook.worksheet_parts()
    table = find_table(package, sheet_entries, name)
    geometry = table.geometry()
    plan = plan_table_fill(geometry, len(adapter.rows), table_name=name)
    sheet = sheets.get(table.sheet.part)
    # The fill rewrites the table area; a shared group that reaches into it is
    # unshared first, so no dependent is left without its master.
    sheet.unshare_range(
        geometry.top_row,
        max(geometry.last_row, plan.new_last_row),
        geometry.first_col,
        geometry.last_col,
    )

    columns = table.columns()
    formula_columns = _calculated_columns(columns, sheet, geometry)
    validate_columns(
        table_columns=[column.name for column in columns],
        data_columns=adapter.columns,
        strategy=config.column_validation_strategy,
        formula_columns=list(formula_columns),
        table_name=name,
        data_rows=len(adapter.rows),
    )
    warn_about_ignored_formula_data(name, adapter.columns, adapter.rows, formula_columns)

    # Captured before anything is cleared: the totals row moves, it is not rewritten.
    totals_row = (
        _snapshot_row(sheet, geometry.last_row, geometry.first_col, geometry.last_col)
        if geometry.totals_rows
        else None
    )
    template_styles = [
        sheet.style(geometry.first_data_row, geometry.first_col + offset)
        for offset in range(len(columns))
    ]

    if plan.overflow_rows is not None:
        _refuse_occupied_growth(package, table, sheet, plan, len(adapter.rows), name)

    # A template saved with a filter applied (or with rows hidden by hand)
    # would keep some of the filled rows out of sight -- and out of a
    # SUBTOTAL(109, …) totals row.  The fill shows everything it writes.
    table.clear_filters()
    sheet.unhide_rows(plan.first_data_row, plan.new_last_row)

    table.set_refs(plan.new_ref, plan.new_autofilter_ref, plan.new_body_ref)

    formula_offsets = {
        offset for offset, column in enumerate(columns) if column.name in formula_columns
    }
    keep_columns = [geometry.first_col + offset for offset in formula_offsets]
    if plan.unused_body_rows is not None:
        _clear_rows(
            sheet,
            plan.unused_body_rows,
            geometry.first_col,
            geometry.last_col,
            keep_columns,
        )
    if plan.vacated_rows is not None:
        _clear_rows(sheet, plan.vacated_rows, geometry.first_col, geometry.last_col)
    if geometry.totals_rows:
        _clear_rows(
            sheet,
            (geometry.last_row, geometry.last_row),
            geometry.first_col,
            geometry.last_col,
        )
    if plan.data_rows == 0:
        _clear_rows(
            sheet,
            (plan.first_data_row, plan.first_data_row),
            geometry.first_col,
            geometry.last_col,
            keep_columns,
        )

    _write_body(
        sheet,
        styles,
        plan=plan,
        columns=columns,
        template_styles=template_styles,
        data_columns=adapter.columns,
        rows=adapter.rows,
        formula_columns=formula_columns,
        config=config,
        table_name=name,
        epoch=_epoch(workbook),
        date1904=workbook.date1904,
    )

    if totals_row is not None:
        _move_row(sheet, totals_row, from_row=geometry.last_row, to_row=plan.new_last_row)
        if plan.vacated_rows is not None:
            _reset_row_style(
                sheet,
                geometry.last_row,
                plan.first_data_row,
                geometry.first_col,
                geometry.last_col,
            )

    if plan.new_body_last_row != geometry.body_last_row:
        _resize_adjacent_formulas(workbook, table, plan, sheets)
        warn_about_untracked_ranges(
            _untracked_range_sources(package, workbook, sheets),
            table_name=name,
            sheet_title=table.sheet.name,
            geometry=geometry,
            new_body_last_row=plan.new_body_last_row,
        )

    if config.extend_conditional_formatting:
        changes = sheet.extend_conditional_formatting(
            table_start_column=geometry.first_col,
            table_end_column=geometry.last_col,
            table_top_row=geometry.top_row,
            first_data_row=plan.first_data_row,
            old_body_last_row=geometry.body_last_row,
            table_last_row=plan.new_body_last_row,
        )
        for old_sqref, new_sqref in changes:
            logger.debug("Extended conditional formatting: %s -> %s", old_sqref, new_sqref)

    sheet.expand_dimension(plan.new_last_row)
    logger.debug("Table '%s' resized: %s -> %s", name, table.ref, plan.new_ref)
    return plan.change_spec(table.sheet.name)


def _calculated_columns(
    columns: Sequence[TableColumnInfo],
    sheet: SheetPart,
    geometry: TableGeometry,
) -> dict[str, str]:
    """Calculated columns of the table, mapped to the formula to write.

    Uses the shared rules: a column is calculated when its definition says so
    *and* the template body agrees (formulas or empty), so a column whose
    definition is stale but which holds values stays a data column.
    """
    formulas: dict[str, str] = {}
    for offset, column in enumerate(columns):
        definition = column.calculated_formula
        if not definition:
            continue
        kinds = [
            _cell_kind(sheet, row, geometry.first_col + offset)
            for row in range(geometry.first_data_row, geometry.body_last_row + 1)
        ]
        evidence = ColumnFormulaEvidence(column.name, definition, True, classify_body(kinds))
        formula = resolve_formula(evidence)
        if formula is not None:
            formulas[column.name] = formula
    return formulas


def _cell_kind(sheet: SheetPart, row: int, column: int) -> Any:  # noqa: ANN401
    """What a template body cell holds: a formula, a value, or nothing."""
    xml = sheet.cell(row, column)
    if xml is None:
        return "empty"
    if sheet.formula(row, column) is not None:
        return "formula"
    return "value" if xmltext.element_text(xml, "v") is not None else "empty"


# ── growth safety ──────────────────────────────────────────────────────────


def _refuse_occupied_growth(
    package: OoxmlPackage,
    table: TablePart,
    sheet: SheetPart,
    plan: TablePlan,
    data_rows: int,
    table_name: str,
) -> None:
    """Refuse to grow into another table, a merged range, data, or a spill."""
    assert plan.overflow_rows is not None
    first_row, last_row = plan.overflow_rows
    geometry = plan.geometry
    regions = [
        BlockedRegion(
            description=f"table {other.display_name!r} at {other.ref}",
            first_row=CellRange(other.ref).min_row,
            last_row=CellRange(other.ref).max_row,
            first_col=CellRange(other.ref).min_col,
            last_col=CellRange(other.ref).max_col,
        )
        for other in tables_on_sheet(package, table.sheet)
        if other.part != table.part
    ]
    regions.extend(
        BlockedRegion(
            description=f"the merged range {merged}",
            first_row=CellRange(merged).min_row,
            last_row=CellRange(merged).max_row,
            first_col=CellRange(merged).min_col,
            last_col=CellRange(merged).max_col,
        )
        for merged in sheet.merged_ranges()
    )
    blocked = blocked_area(
        (first_row, last_row),
        geometry.first_col,
        geometry.last_col,
        regions,
    )
    if blocked is not None:
        msg = (
            f"Table '{table_name}' cannot expand to {data_rows} rows: "
            f"{blocked.description} is in the way. Move or remove the content below "
            "the table first."
        )
        raise ValueError(msg)

    for row in range(first_row, last_row + 1):
        for column in range(geometry.first_col, geometry.last_col + 1):
            reference = f"{get_column_letter(column)}{row}"
            spill = sheet.array_range_at(row, column)
            if spill is not None:
                msg = (
                    f"Table '{table_name}' cannot expand to {data_rows} rows: "
                    f"cell {reference} is part of the array formula {spill}. "
                    "Move or remove the content below the table first."
                )
                raise ValueError(msg)
            if sheet.has_value(row, column):
                current = _display_value(sheet, row, column)
                msg = (
                    f"Table '{table_name}' cannot expand to {data_rows} rows: "
                    f"cell {reference} already contains data ({current}). Move or "
                    "remove the content below the table first."
                )
                raise ValueError(msg)


def _display_value(sheet: SheetPart, row: int, column: int) -> str:
    """A short description of what a cell holds, for an error message."""
    xml = sheet.cell(row, column)
    if xml is None:
        return "''"
    formula = sheet.formula(row, column)
    if formula is not None:
        return repr(formula)
    text = xmltext.element_text(xml, "v")
    if xmltext.attribute(xml, "t") == "s" and text is not None and text.isdigit():
        shared = sheet.package.shared_strings()
        index = int(text)
        if 0 <= index < len(shared):
            return repr(shared[index])
    if text is None:
        inline = xmltext.element_text(xml, "t")
        return repr(inline) if inline is not None else "''"
    try:
        return repr(int(text))
    except ValueError:
        pass
    try:
        return repr(float(text))
    except ValueError:
        return repr(text)


# ── cell writing ───────────────────────────────────────────────────────────


def _cell_style(
    sheet: SheetPart,
    row: int,
    column: int,
    template_style: str | None,
    config: FillXLConfig,
) -> str | None:
    """The style a written cell gets.

    With ``copy_table_styles`` on (the default) new cells look like the
    template's first data row, which is what Excel does when a table grows.
    With it off the engine must not touch styling at all: an existing cell
    keeps whatever it had, and a new one gets the default style.
    """
    if config.copy_table_styles:
        return template_style
    return sheet.style(row, column)


def _refuse_spill_write(sheet: SheetPart, row: int, column: int) -> None:
    """Raise when a write would land inside an array (spilled) formula range."""
    spill = sheet.array_range_at(row, column)
    if spill is None:
        return
    msg = (
        f"Cannot write cell {sheet.reference(row, column)}: it is part of the array "
        f"formula {spill}. Change or remove that formula first."
    )
    raise ValueError(msg)


def _write(
    sheet: SheetPart,
    styles: StylesPart,
    row: int,
    column: int,
    value: Any,  # noqa: ANN401
    style: str | None,
    *,
    escape_formulas: bool,
    epoch: dt.datetime,
) -> None:
    """Write one cell value with the style a template cell gave it."""
    _refuse_spill_write(sheet, row, column)
    sheet.write_cell(
        row,
        column,
        build_cell(
            sheet.reference(row, column),
            value,
            style=_date_style(styles, style, value),
            escape_formulas=escape_formulas,
            epoch=epoch,
        ),
    )


def _clear_cell(sheet: SheetPart, row: int, column: int) -> None:
    """Clear a cell's value but keep its style, as the openpyxl engine does."""
    if sheet.cell(row, column) is None:
        return
    _refuse_spill_write(sheet, row, column)
    style = sheet.style(row, column)
    if style is None:
        sheet.delete_cell(row, column)
        return
    sheet.write_cell(row, column, f'<c r="{sheet.reference(row, column)}" s="{style}"/>')


def _clear_rows(
    sheet: SheetPart,
    rows: tuple[int, int],
    first_col: int,
    last_col: int,
    keep: Sequence[int] = (),
) -> None:
    """Clear values (not styles) in *rows* between *first_col* and *last_col*."""
    first_row, last_row = rows
    skipped = set(keep)
    for row in range(first_row, last_row + 1):
        for column in range(first_col, last_col + 1):
            if column not in skipped:
                _clear_cell(sheet, row, column)


def _snapshot_row(sheet: SheetPart, row: int, first_col: int, last_col: int) -> list[str | None]:
    """Capture the ``<c>`` elements of a row, so it can be moved later."""
    return [sheet.cell(row, column) for column in range(first_col, last_col + 1)]


def _move_row(sheet: SheetPart, cells: list[str | None], *, from_row: int, to_row: int) -> None:
    """Write snapped *cells* on *to_row*, shifting array refs by the row delta."""
    delta = to_row - from_row
    for xml in cells:
        if xml is None:
            continue
        reference = xmltext.attribute(xml, "r")
        if reference is None:  # pragma: no cover - Excel always writes r
            continue
        column = _column_of(reference)
        moved = xmltext.set_tag_attribute(xml, "r", sheet.reference(to_row, column))
        spill = array_range(moved)
        if spill is not None:
            moved = moved.replace(
                f'ref="{spill}"',
                f'ref="{_shift_range(spill, delta)}"',
                1,
            )
        sheet.write_cell(to_row, column, moved)


def _column_of(reference: str) -> int:
    """The column number of an A1 reference."""
    _, column = split_reference(reference)
    return column


def _shift_range(reference: str, delta: int) -> str:
    """*reference* moved *delta* rows down."""
    bounds = CellRange(reference)
    return CellRange(
        min_col=bounds.min_col,
        min_row=bounds.min_row + delta,
        max_col=bounds.max_col,
        max_row=bounds.max_row + delta,
    ).coord


def _reset_row_style(
    sheet: SheetPart,
    row: int,
    template_row: int,
    first_col: int,
    last_col: int,
) -> None:
    """Give *row* the style of *template_row*, so no leftover formatting remains."""
    if row == template_row:
        return
    for column in range(first_col, last_col + 1):
        if sheet.cell(row, column) is None:
            continue
        source = sheet.style(template_row, column)
        if source is None:
            continue
        sheet.set_style(row, column, source)


def _write_body(
    sheet: SheetPart,
    styles: StylesPart,
    *,
    plan: TablePlan,
    columns: Sequence[TableColumnInfo],
    template_styles: list[str | None],
    data_columns: list[str],
    rows: Sequence[Sequence[Any]],
    formula_columns: dict[str, str],
    config: FillXLConfig,
    table_name: str,
    epoch: dt.datetime,
    date1904: bool,
) -> None:
    """Write the data body: calculated columns, data columns, cleared leftovers."""
    data_column_map = {str(name): index for index, name in enumerate(data_columns)}
    geometry = plan.geometry
    warned: set[tuple[str | None, str | None]] = set()
    for offset, column in enumerate(columns):
        column_name = column.name
        column_number = geometry.first_col + offset
        template_style = template_styles[offset] if offset < len(template_styles) else None
        if column_name in formula_columns:
            formula = formula_columns[column_name]
            for row_index in range(plan.data_rows):
                row = geometry.first_data_row + row_index
                # Excel copies the definition down, which moves its relative
                # row references; the same text on every row would make every
                # row compute the first row's result.
                _write(
                    sheet,
                    styles,
                    row,
                    column_number,
                    shift_relative_rows(formula, row_index),
                    _cell_style(sheet, row, column_number, template_style, config),
                    # A calculated column is template content, not user data:
                    # escape_formulas must not turn its definition into text.
                    escape_formulas=False,
                    epoch=epoch,
                )
        elif column_name in data_column_map:
            data_index = data_column_map[column_name]
            for row_index, row_values in enumerate(rows):
                raw = row_values[data_index]
                value = prepare_table_cell(
                    raw,
                    table=table_name,
                    row=row_index + 1,
                    column=column_name,
                    warned=warned,
                    date1904=date1904,
                )
                row = geometry.first_data_row + row_index
                _write(
                    sheet,
                    styles,
                    row,
                    column_number,
                    value,
                    _cell_style(sheet, row, column_number, template_style, config),
                    escape_formulas=config.escape_formulas,
                    epoch=epoch,
                )
        else:
            # The data has no such column: clear the template's sample values.
            _clear_rows(
                sheet,
                (plan.first_data_row, plan.new_body_last_row),
                column_number,
                column_number,
            )


def _resize_adjacent_formulas(
    workbook: WorkbookPart,
    table: TablePart,
    plan: TablePlan,
    sheets: SheetStore,
) -> None:
    """Resize formulas that tracked the table body, as Excel does.

    Excel rewrites ``=SUM(B4:B4)`` to ``=SUM(B4:B7)`` when the table grows, on
    the table's sheet and on any other sheet that refers to it.

    The table's data body is skipped: those formulas belong to the table.  The
    totals row is *not* skipped -- it moved with the table, and an A1 formula
    there has to follow the body too.
    """
    geometry = plan.geometry
    table_sheet = table.sheet.name
    parts = [entry.part for entry in workbook.worksheet_parts()]
    for sheet in sheets.formula_parts(parts):
        own_sheet = sheet.part == table.sheet.part
        for row, column, formula in list(sheet.formula_cells()):
            xml = sheet.cell(row, column)
            if xml is not None and array_range(xml) is not None:
                # openpyxl leaves array formulas alone; rewriting half of one
                # would desynchronise the ref Excel validates.
                continue
            in_body = (
                own_sheet
                and plan.first_data_row <= row <= plan.new_body_last_row
                and geometry.first_col <= column <= geometry.last_col
            )
            if in_body:
                continue
            formula_sheet = _sheet_title(workbook, sheet.part)
            if not may_track_table(
                formula,
                table_sheet=table_sheet,
                formula_sheet=formula_sheet,
            ):
                continue
            try:
                updated = extend_table_ranges(
                    formula,
                    sheet_title=table_sheet,
                    formula_sheet=formula_sheet,
                    first_col=geometry.first_col,
                    last_col=geometry.last_col,
                    old_body=(geometry.first_data_row, geometry.body_last_row),
                    new_body=(plan.first_data_row, plan.new_body_last_row),
                )
            except TokenizerError:
                # A formula openpyxl cannot read is not one this fill can
                # rewrite; failing the whole fill over it would be worse.
                logger.debug(
                    "Cannot read the formula in %s!%s; leaving it alone",
                    formula_sheet,
                    sheet.reference(row, column),
                )
                continue
            if updated is not None:
                logger.debug("Resized formula %s -> %s", formula, updated)
                sheet.set_formula(row, column, updated)


def _sheet_title(workbook: WorkbookPart, part: str) -> str:
    """The tab name of a sheet part."""
    for entry in workbook.sheets():
        if entry.part == part:
            return entry.name
    return ""  # pragma: no cover - every part comes from the workbook's sheets


def _untracked_range_sources(
    package: OoxmlPackage,
    workbook: WorkbookPart,
    sheets: SheetStore,
) -> list[RangeSource]:
    """Every defined name, pivot source and validation list that holds a range."""
    sources: list[RangeSource] = [
        RangeSource(f"the defined name {entry.name!r}", sheet_title, reference)
        for entry in workbook.defined_names()
        for sheet_title, reference in destination_references(entry.text)
    ]
    sources.extend(
        RangeSource("a pivot cache source", sheet_title, reference)
        for sheet_title, reference in pivot_worksheet_sources(package)
    )
    for entry in workbook.worksheet_parts():
        sources.extend(_validation_sources(sheets.get(entry.part), entry.name))
    return sources


def _validation_sources(sheet: SheetPart, sheet_title: str) -> list[RangeSource]:
    """The range a sheet's data validations read their lists from.

    An unqualified list source means the sheet the validation sits on.
    """
    sources: list[RangeSource] = []
    for match in _DATA_VALIDATION.finditer(sheet.tail):
        element = match.group(0)
        sqref = xmltext.attribute(element, "sqref")
        description = f"the data validation over {sqref}" if sqref else "a data validation"
        for formula in _validation_formulas(element):
            parsed = sole_reference(formula)
            if parsed is None:
                continue
            target_sheet, reference = parsed
            sources.append(RangeSource(description, target_sheet or sheet_title, reference))
    return sources


def _validation_formulas(element: str) -> list[str]:
    """The formulas of one ``<dataValidation>``, including the x14 spelling.

    Excel stores a cross-sheet list source added by the newer UI as
    ``<x14:formula1><xm:f>…</xm:f></x14:formula1>`` inside an ``extLst``;
    the classic part keeps the reference as the element's own text.
    """
    formulas: list[str] = []
    for name in ("formula1", "formula2"):
        match = xmltext.element_re(name).search(element)
        if match is None:
            continue
        raw = match.group(0)
        child = xmltext.element_re("f").search(raw)
        text = (
            xmltext.element_text(child.group(0), "f")
            if child is not None
            else xmltext.element_text(raw, name)
        )
        if text is not None:
            formulas.append(text)
    return formulas
