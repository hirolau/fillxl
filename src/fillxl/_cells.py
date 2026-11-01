"""Named cell population logic."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from fillxl._normalize import (
    check_cell_value,
    check_named_cell_scalar,
    check_nonfinite_value,
    is_text_cell_value,
    prepare_cell_value,
)
from fillxl._warnings import warn_user

if TYPE_CHECKING:
    from collections.abc import Sequence

    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.workbook.workbook import Workbook
    from openpyxl.worksheet.table import Table

logger = logging.getLogger(__name__)


def named_cell_target(
    name: str,
    destinations: Sequence[tuple[str, str]],
) -> tuple[str, str] | None:
    """Pick the single cell a scalar named-cell value is written to (both backends).

    Returns:
        ``(sheet_name, address)``, or ``None`` when the name has no destination.

    Raises:
        ValueError: If the first destination is a multi-cell range.

    """
    if not destinations:
        warn_user(f"Named cell {name!r} has no destinations, so nothing was written")
        return None
    sheet_name, address = destinations[0]
    if ":" in address:
        msg = (
            f"Named cell '{name}' refers to a multi-cell range "
            f"({sheet_name}!{address}), but only single cells "
            f"can be filled with a scalar value"
        )
        raise ValueError(msg)
    if len(destinations) > 1:
        warn_user(
            f"Named cell {name!r} has {len(destinations)} destinations; only the first "
            f"one ({sheet_name}!{address}) was updated",
        )
    return sheet_name.replace("''", "'"), address


def find_defined_name(workbook: Workbook, name: str) -> DefinedName | None:
    """Find a defined name, including names scoped to a single sheet.

    ``Workbook.defined_names`` only holds workbook-global names; names scoped
    to a sheet live on that sheet (``Worksheet.defined_names``).  Looking in
    both keeps the openpyxl backend in step with win32com, which resolves
    sheet-scoped names through ``Workbook.Names``.

    Names are matched case-insensitively after an exact match has failed:
    Excel and the COM object model both resolve ``single`` to ``Single``, so
    the openpyxl engine must not raise where the other engine writes.

    Args:
        workbook: openpyxl Workbook instance
        name: Defined name to look up

    Returns:
        The :class:`~openpyxl.workbook.defined_name.DefinedName`, or ``None``.

    """
    if name in workbook.defined_names:
        return workbook.defined_names[name]

    for worksheet in workbook.worksheets:
        scoped_names = worksheet.defined_names
        if name in scoped_names:
            return scoped_names[name]

    return _find_defined_name_ignoring_case(workbook, name)


def _find_defined_name_ignoring_case(workbook: Workbook, name: str) -> DefinedName | None:
    """The only name that differs from *name* by case, if there is exactly one.

    Excel forbids two names that differ only by case in the same scope, so a
    match here is unambiguous for global names.  A global name and a
    sheet-scoped name can still differ by case; the global one wins, matching
    the order of the exact-match lookup.
    """
    folded = name.casefold()
    for collection in [workbook.defined_names, *(ws.defined_names for ws in workbook.worksheets)]:
        matches = [candidate for candidate in collection if candidate.casefold() == folded]
        if len(matches) == 1:
            return collection[matches[0]]
        if matches:
            return None
    return None


def table_header_conflict(worksheet: Worksheet, address: str) -> Table | None:
    """The table whose header row *address* points into, or ``None``.

    Writing a header cell leaves the table's ``tableColumns`` name behind, and
    a workbook where the two disagree makes Excel repair (or refuse) the file.

    Examples:
        >>> import openpyxl
        >>> from openpyxl.worksheet.table import Table, TableColumn
        >>> wb = openpyxl.Workbook()
        >>> ws = wb.active
        >>> ws.add_table(Table(
        ...     displayName="T",
        ...     ref="A1:B2",
        ...     tableColumns=[TableColumn(id=1, name="a"), TableColumn(id=2, name="b")],
        ... ))
        >>> table_header_conflict(ws, "A1").name
        'T'
        >>> table_header_conflict(ws, "A2") is None
        True
        >>> table_header_conflict(ws, "C1") is None
        True

    """
    try:
        target = CellRange(address)
    except ValueError:
        return None
    for table in worksheet.tables.values():
        if getattr(table, "headerRowCount", 1) == 0:
            continue
        header = CellRange(table.ref)
        if header.min_row == target.min_row and header.min_col <= target.min_col <= header.max_col:
            return table
    return None


def header_write_error(sheet_name: str, address: str, table_name: str) -> ValueError:
    """The error for a named-cell write that would corrupt a table's header row.

    Both backends raise it, so a template fails the same way whichever engine
    runs it.  Renaming a header through a name would either leave the table's
    column name out of step with the cell (openpyxl) or silently break the
    data-to-column mapping the fill is about to use (Excel).

    Examples:
        >>> error = header_write_error("Report", "$A$1", "T")
        >>> str(error).split(",")[0]
        "Report!$A$1 is a header cell of table 'T'"

    """
    msg = (
        f"{sheet_name}!{address} is a header cell of table {table_name!r}, so writing it "
        "would leave the table's column name out of step with its cell. Rename the column "
        "in Excel instead, or point the name at another cell."
    )
    return ValueError(msg)


def populate_named_cells(
    workbook: Workbook,
    cells: dict[str, Any],
    *,
    allow_missing: bool = False,
    escape_formulas: bool = False,
) -> None:
    """Populate named cells with scalar values.

    Args:
        workbook: openpyxl Workbook instance
        cells: Dictionary mapping named cell names to values
        allow_missing: If False, raise error if named cell is not found
        escape_formulas: If True, strings starting with ``=`` are written as
            literal text instead of becoming formulas.

    Raises:
        KeyError: If a named cell is not found and allow_missing is False
        ValueError: If the name points at a table header row, which cannot be
            renamed without desynchronising the table's column.

    """
    for name, value in cells.items():
        defined_name = find_defined_name(workbook, name)
        if defined_name is None:
            if allow_missing:
                warn_user(f"Named cell {name!r} not found in workbook, nothing was written")
                continue
            msg = f"Named cell '{name}' not found in workbook"
            raise KeyError(msg)

        # Refuse containers before normalisation: str(value) would otherwise
        # write a DataFrame or list as its repr.
        check_named_cell_scalar(name, value)
        # An infinity is written as an empty cell; say so instead of losing it.
        check_nonfinite_value(value, name=name)

        # Normalise value for Excel compatibility
        normalised_value = prepare_cell_value(value)
        check_cell_value(normalised_value, name=name)

        target = named_cell_target(name, list(defined_name.destinations))
        if target is None:
            continue
        worksheet_name, cell_address = target

        worksheet = workbook[worksheet_name]
        if not isinstance(worksheet, Worksheet):
            logger.warning("Worksheet '%s' is not a regular worksheet, skipping", worksheet_name)
            continue

        conflict = table_header_conflict(worksheet, cell_address)
        if conflict is not None:
            raise header_write_error(worksheet_name, cell_address, str(conflict.name))

        cell = worksheet[cell_address]
        cell.value = normalised_value

        # Force text for strings openpyxl would reinterpret: error values
        # such as "#N/A", and formulas when escaping is enabled.
        if is_text_cell_value(normalised_value, escape_formulas=escape_formulas):
            cell.data_type = "s"

        logger.debug(
            "Set named cell '%s' (%s!%s) = %s",
            name,
            worksheet_name,
            cell_address,
            normalised_value,
        )
