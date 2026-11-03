"""Calculated-column detection and formula row arithmetic, shared by the backends.

A table column is a calculated column when its definition carries a
``calculatedColumnFormula`` and the template body agrees (it holds formulas or
is empty).  Both backends apply the same rule, so they keep the same columns.

openpyxl writes formulas as text, so it has to do the row arithmetic itself:
Excel adjusts relative references per row when it fills a calculated column,
and it resizes ranges that track a table when the table grows or shrinks.
:func:`shift_relative_rows` and :func:`extend_table_ranges` reproduce those
two rules on the formula text.
"""

from __future__ import annotations

import logging
import math
import re
import sys
from typing import TYPE_CHECKING, Any, Literal, NamedTuple

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

from fillxl._utils import table_column_name
from fillxl._warnings import warn_user

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping, Sequence

    from openpyxl.worksheet.table import Table
    from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

BodyState = Literal["formula", "mixed", "empty", "values"]
CellKind = Literal["formula", "empty", "value"]

_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384
# A single A1 reference, optionally qualified by a sheet name.  A defined name
# such as ``Rate`` has no digits and a structured reference contains ``[``, so
# neither matches; ``LOG10(`` is a function token, never an operand.
# The shape alone is not enough: ``YTD2024`` looks like column ``YTD`` and row
# 2024, but that column is past ``XFD``, so Excel reads it as a defined name.
# Every match therefore has to pass :func:`_is_a1_operand`.
_A1_OPERAND = re.compile(
    r"^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$"
)
_TWO_CELL_RANGE = re.compile(
    r"^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?"
    r"(?P<first>\$?[A-Za-z]{1,3}\$?\d+):(?P<last>\$?[A-Za-z]{1,3}\$?\d+)$"
)
_COLUMN_ROW = re.compile(r"(?P<col>\$?[A-Za-z]{1,3})(?P<row>\$?\d+)")


def _is_a1_operand(operand: str) -> bool:
    """Whether *operand* is an A1 reference inside Excel's grid.

    A defined name whose letters look like a column past ``XFD`` is spelled the
    same way as a reference (``YTD2024``, ``ZAR1``), and Excel never shifts it
    when a formula is copied.  Treating it as a reference would silently point
    the copied formula at another name — or at nothing.

    Examples:
        >>> _is_a1_operand("B4")
        True
        >>> _is_a1_operand("Data!$B$4:$C$9")
        True
        >>> _is_a1_operand("XFD1048576")
        True
        >>> _is_a1_operand("YTD2024")
        False
        >>> _is_a1_operand("ZAR1")
        False
        >>> _is_a1_operand("Rate")
        False

    """
    if _A1_OPERAND.match(operand) is None:
        return False
    _, refs = _split_sheet(operand)
    return all(_endpoint_in_bounds(reference) for reference in refs.split(":"))


def _endpoint_in_bounds(reference: str) -> bool:
    """Whether one ``$?COL$?ROW`` reference lands inside the sheet."""
    match = _COLUMN_ROW.fullmatch(reference)
    if match is None:  # pragma: no cover - _A1_OPERAND guarantees the shape
        return False
    return (
        1 <= _column_index(match.group("col")) <= _MAX_COLUMN
        and 1 <= int(match.group("row").lstrip("$")) <= _MAX_ROW
    )


def a1_reference(operand: str) -> tuple[str | None, str] | None:
    """``(sheet_title, reference)`` when *operand* is one A1 reference.

    The sheet title is unquoted (Excel doubles a quote inside a quoted name),
    and ``None`` when the reference is unqualified — it then means the sheet the
    reference sits on.

    Whole-row and whole-column references (``$1:$1``, ``$A:$A``) are not
    rectangles this package can compare, so they return ``None``.

    Examples:
        >>> a1_reference("Data!$B$4:$C$9")
        ('Data', '$B$4:$C$9')
        >>> a1_reference("'Bob''s Data'!$A$1")
        ("Bob's Data", '$A$1')
        >>> a1_reference("$A$2:$B$2")
        (None, '$A$2:$B$2')
        >>> a1_reference("$1:$1") is None
        True
        >>> a1_reference("$A:$A") is None
        True
        >>> a1_reference("YTD2024") is None
        True

    """
    if not _is_a1_operand(operand):
        return None
    prefix, refs = _split_sheet(operand)
    if not prefix:
        return None, refs
    title = prefix[:-1]
    if title.startswith("'") and title.endswith("'"):
        title = title[1:-1].replace("''", "'")
    return title, refs


def sole_reference(text: str) -> tuple[str | None, str] | None:
    """``(sheet_title, reference)`` when *text* is nothing but one A1 reference.

    A data-validation list source and a defined name are written as a bare
    reference; a custom validation rule (``A1>0``), a literal list
    (``"a,b,c"``) and a formula (``OFFSET(…)``) are not, and are refused here
    rather than being parsed into a range that means something else.

    Examples:
        >>> sole_reference("$H$2:$H$10")
        (None, '$H$2:$H$10')
        >>> sole_reference("=Data!$A$2:$A$5")
        ('Data', '$A$2:$A$5')
        >>> sole_reference("A1>0") is None
        True
        >>> sole_reference('"a,b,c"') is None
        True
        >>> sole_reference("OFFSET(Data!$A$1,0,0)") is None
        True
        >>> sole_reference("Data!$A$2:$A$5,Other!$B$1:$B$3") is None
        True

    """
    stripped = text.strip().removeprefix("=").strip()
    if not stripped or stripped.startswith('"'):
        return None
    try:
        tokens = Tokenizer(f"={stripped}").items
    except TokenizerError:
        # A hand-written or damaged definition is not a reference this can use;
        # it must not turn a fill into a crash.
        return None
    operands = [token.value for token in tokens if token.value.strip()]
    if len(operands) != 1:
        return None
    return a1_reference(operands[0])


def _split_sheet(operand: str) -> tuple[str, str]:
    """``("Data!", "B4:B4")`` for an operand that may carry a sheet name.

    Examples:
        >>> _split_sheet("Data!B4:B4")
        ('Data!', 'B4:B4')
        >>> _split_sheet("'My Sheet'!B4")
        ("'My Sheet'!", 'B4')
        >>> _split_sheet("B4")
        ('', 'B4')

    """
    prefix, separator, refs = operand.rpartition("!")
    return (prefix + separator, refs) if separator else ("", operand)


def _sheet_matches(prefix: str, sheet_title: str) -> bool:
    """Whether an operand's sheet prefix names *sheet_title*.

    Examples:
        >>> _sheet_matches("Data!", "Data")
        True
        >>> _sheet_matches("Data!", "Other")
        False
        >>> _sheet_matches("'My Sheet'!", "My Sheet")
        True
        >>> _sheet_matches("'It''s'!", "It's")
        True
        >>> _sheet_matches("", "Data")
        False

    """
    if not prefix:
        return False
    name = prefix[:-1]
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return name.casefold() == sheet_title.casefold()


def _replace_rows(refs: str, first_row: int, last_row: int) -> str:
    """Set the rows of the endpoints of ``refs`` (``B4:B7``), keeping ``$`` markers."""
    match = _TWO_CELL_RANGE.match(refs)
    if match is None:
        return refs
    first = _with_row(match.group("first"), first_row)
    last = _with_row(match.group("last"), last_row)
    return f"{first}:{last}"


def _with_row(reference: str, row: int) -> str:
    """*reference* with its row replaced by *row*, keeping any ``$``.

    Examples:
        >>> _with_row("B4", 7)
        'B7'
        >>> _with_row("$B$4", 7)
        '$B$7'
        >>> _with_row("B$4", 7)
        'B$7'

    """
    match = _COLUMN_ROW.fullmatch(reference)
    if match is None:
        return reference
    row_text = match.group("row")
    dollar = "$" if row_text.startswith("$") else ""
    return f"{match.group('col')}{dollar}{row}"


def _shift_operand(operand: str, row_offset: int, column_offset: int) -> str:
    """Move every relative reference of *operand* by the given offsets."""
    prefix, refs = _split_sheet(operand)

    def replace(match: re.Match[str]) -> str:
        column_text = match.group("col")
        row_text = match.group("row")
        column = column_text.lstrip("$")
        row = row_text.lstrip("$")
        if row_offset and not row_text.startswith("$"):
            new_row = int(row) + row_offset
            if not 1 <= new_row <= _MAX_ROW:
                msg = (
                    f"Filling this table would move the reference {operand!r} to row "
                    f"{new_row}, outside the sheet"
                )
                raise ValueError(msg)
            row = str(new_row)
        if column_offset and not column_text.startswith("$"):
            new_column = _column_index(column) + column_offset
            if not 1 <= new_column <= _MAX_COLUMN:
                msg = (
                    f"Filling this table would move the reference {operand!r} to "
                    f"column {new_column}, outside the sheet"
                )
                raise ValueError(msg)
            column = get_column_letter(new_column)
        dollar_column = "$" if column_text.startswith("$") else ""
        dollar_row = "$" if row_text.startswith("$") else ""
        return f"{dollar_column}{column}{dollar_row}{row}"

    return prefix + _COLUMN_ROW.sub(replace, refs)


def shift_relative(formula: str, *, rows: int = 0, columns: int = 0) -> str:
    """*formula* with its relative A1 references moved by *rows* and *columns*.

    Excel moves relative references when a formula is copied, including the
    column part when it is filled right — which is how a shared formula group
    is stored: one text, many positions.

    Examples:
        >>> shift_relative("=B5*C5", rows=2)
        '=B7*C7'
        >>> shift_relative("=SUM(B5:B5)", columns=2)
        '=SUM(D5:D5)'
        >>> shift_relative("=B5*$C5", rows=1, columns=1)
        '=C6*$C6'
        >>> shift_relative("=[@Qty]*2", rows=3, columns=1)
        '=[@Qty]*2'

    """
    if rows == 0 and columns == 0:
        return formula
    tokenizer = Tokenizer(formula)
    parts: list[str] = []
    changed = False
    for token in tokenizer.items:
        if token.type == "OPERAND" and token.subtype == "RANGE" and _is_a1_operand(token.value):
            shifted = _shift_operand(token.value, rows, columns)
            changed = changed or shifted != token.value
            parts.append(shifted)
        else:
            parts.append(token.value)
    if not changed:
        return formula
    joined = "".join(parts)
    return f"={joined}" if formula.startswith("=") else joined


def shift_relative_rows(formula: str, offset: int) -> str:
    """*formula* with its relative A1 rows moved down by *offset*.

    Excel fills a calculated column by copying the definition down, which
    moves relative row references and leaves ``$``-anchored rows alone.
    Structured references (``[@Qty]``) and string literals are row-relative
    (or not references at all) and are never touched.

    Examples:
        >>> shift_relative_rows("=B5*C5", 0)
        '=B5*C5'
        >>> shift_relative_rows("=B5*C5", 2)
        '=B7*C7'
        >>> shift_relative_rows("=B5*$C$5", 1)
        '=B6*$C$5'
        >>> shift_relative_rows("=B5*C$5", 1)
        '=B6*C$5'
        >>> shift_relative_rows("=[@Qty]*[@Price]", 3)
        '=[@Qty]*[@Price]'
        >>> shift_relative_rows('=IF(A5>0,"B5",B5)', 1)
        '=IF(A6>0,"B5",B6)'
        >>> shift_relative_rows("=Data!B5*2", 1)
        '=Data!B6*2'
        >>> shift_relative_rows("=LOG10(B5)", 1)
        '=LOG10(B6)'
        >>> shift_relative_rows("=YTD2024*2", 1)
        '=YTD2024*2'
        >>> shift_relative_rows("=ZAR1*2", 1)
        '=ZAR1*2'

    """
    return shift_relative(formula, rows=offset)


def extend_table_ranges(
    formula: str,
    *,
    sheet_title: str,
    first_col: int,
    last_col: int,
    old_body: tuple[int, int],
    new_body: tuple[int, int],
    formula_sheet: str | None = None,
) -> str | None:
    """*formula* with ranges that tracked the table body resized, or ``None``.

    Excel resizes a range when it covers exactly one of the table's column
    runs over exactly the body rows -- ``=SUM(B4:B4)`` becomes ``=SUM(B4:B7)``
    when the table grows.  Single-cell references are not touched, and neither
    are ranges on another sheet, ranges that start or end outside the body, or
    string literals that merely look like one.

    An unqualified reference means the sheet the formula lives on, so
    *formula_sheet* (default: the table's sheet) decides whether it points at
    the table.

    Examples:
        >>> kw = dict(sheet_title="Data", first_col=2, last_col=2,
        ...           old_body=(4, 4), new_body=(4, 7))
        >>> extend_table_ranges("=SUM(B4:B4)", **kw)
        '=SUM(B4:B7)'
        >>> extend_table_ranges("=SUM($B$4:$B$4)", **kw)
        '=SUM($B$4:$B$7)'
        >>> extend_table_ranges("=SUM(Data!B4:B4)", **kw)
        '=SUM(Data!B4:B7)'
        >>> extend_table_ranges("=SUM(Data!B4:B4)", formula_sheet="Other", **kw)
        '=SUM(Data!B4:B7)'
        >>> extend_table_ranges("=SUM(B4:B4)", formula_sheet="Other", **kw) is None
        True
        >>> extend_table_ranges("=SUM(B4:B5)", **kw) is None
        True
        >>> extend_table_ranges("=SUM(Z1:Z1)", **kw) is None
        True
        >>> extend_table_ranges("=SUM(YTD2024:YTD2024)", **kw) is None
        True
        >>> extend_table_ranges("=B4*2", **kw) is None
        True
        >>> extend_table_ranges('=IF(A1>0,"B4:B4",SUM(B4:B4))', **kw)
        '=IF(A1>0,"B4:B4",SUM(B4:B7))'

    """
    tokenizer = Tokenizer(formula)
    parts: list[str] = []
    changed = False
    for token in tokenizer.items:
        value = token.value
        if (
            token.type == "OPERAND"
            and token.subtype == "RANGE"
            and _TWO_CELL_RANGE.match(value)
            and _is_a1_operand(value)
        ):
            prefix, refs = _split_sheet(value)
            points_at_table = (
                _sheet_matches(prefix, sheet_title)
                if prefix
                else (formula_sheet or sheet_title).casefold() == sheet_title.casefold()
            )
            if points_at_table and _tracks_body(
                refs,
                first_col=first_col,
                last_col=last_col,
                old_body=old_body,
                new_body=new_body,
            ):
                resized = _replace_rows(refs, new_body[0], new_body[1])
                changed = changed or resized != refs
                parts.append(prefix + resized)
                continue
        parts.append(value)
    if not changed:
        return None
    joined = "".join(parts)
    return f"={joined}" if formula.startswith("=") else joined


def _tracks_body(
    refs: str,
    *,
    first_col: int,
    last_col: int,
    old_body: tuple[int, int],
    new_body: tuple[int, int],
) -> bool:
    """Whether the range ``refs`` covers a table column run over the whole body."""
    if not _is_a1_operand(refs):
        return False
    match = _TWO_CELL_RANGE.match(refs)
    if match is None:
        return False
    rows: list[int] = []
    cols: list[int] = []
    for reference in (match.group("first"), match.group("last")):
        found = _COLUMN_ROW.fullmatch(reference)
        if found is None:
            return False
        rows.append(int(found.group("row").lstrip("$")))
        cols.append(_column_index(found.group("col")))
    first_row, last_row = min(rows), max(rows)
    first_range_col, last_range_col = min(cols), max(cols)
    return (
        (first_row, last_row) == old_body
        and first_range_col >= first_col
        and last_range_col <= last_col
        and (first_row, last_row) != new_body
    )


def _column_index(column: str) -> int:
    """The 1-based number of an Excel column name, ignoring a ``$``.

    Examples:
        >>> _column_index("$A")
        1
        >>> _column_index("B")
        2
        >>> _column_index("AA")
        27

    """
    index = 0
    for character in column.lstrip("$").upper():
        index = index * 26 + ord(character) - ord("A") + 1
    return index


class ColumnFormulaEvidence(NamedTuple):
    """What a backend observed about one table column before writing to it.

    Attributes:
        name: Column name.
        definition: ``calculatedColumnFormula`` text without ``=``, or ``None``.
        definitions_known: ``False`` when the file format hides definitions
            (``.xlsb``/``.xls`` read through COM).
        body_state: What the template body cells of this column hold.
        body_formula: Formula of the first body cell including ``=`` (COM only;
            openpyxl passes ``None`` and writes the definition instead).

    """

    name: str
    definition: str | None
    definitions_known: bool
    body_state: BodyState
    body_formula: str | None = None


def formula_definitions(table: Table) -> dict[str, str]:
    """Map each column that has a ``calculatedColumnFormula`` to its text (no ``=``).

    Constants such as ``100`` are kept: ``=100`` is a valid calculated column.
    Whether a definition is still live is decided by :func:`resolve_formula`.
    """
    definitions: dict[str, str] = {}
    for column in table.tableColumns:
        formula = column.calculatedColumnFormula
        text = getattr(formula, "attr_text", None) if formula is not None else None
        if text:
            definitions[table_column_name(column)] = str(text)
    return definitions


def classify_body(kinds: Iterable[CellKind]) -> BodyState:
    """Summarise the kinds of the body cells of one column.

    Examples:
        >>> classify_body(["formula", "formula"])
        'formula'
        >>> classify_body(["formula", "empty"])
        'formula'
        >>> classify_body(["formula", "value"])
        'mixed'
        >>> classify_body(["empty", "empty"])
        'empty'
        >>> classify_body(["value", "empty"])
        'values'

    """
    seen = set(kinds)
    if "formula" in seen:
        return "mixed" if "value" in seen else "formula"
    return "values" if "value" in seen else "empty"


def resolve_formula(evidence: ColumnFormulaEvidence) -> str | None:
    """Return the formula a column must keep (with ``=``), or ``None`` for a data column.

    Examples:
        >>> E = ColumnFormulaEvidence
        >>> resolve_formula(E("C", "[@A]*2", True, "empty"))
        '=[@A]*2'
        >>> resolve_formula(E("C", "[@A]*2", True, "formula", "=[@A]*2"))
        '=[@A]*2'
        >>> resolve_formula(E("C", "100", True, "formula"))
        '=100'
        >>> resolve_formula(E("C", "100", True, "values")) is None
        True
        >>> resolve_formula(E("C", None, True, "formula")) is None
        True
        >>> resolve_formula(E("C", None, False, "formula", "=A2*2"))
        '=A2*2'
        >>> resolve_formula(E("C", None, False, "mixed", "=A2*2")) is None
        True

    """
    if not evidence.definitions_known:
        # Without definitions only a column that Excel filled completely is trusted.
        return evidence.body_formula if evidence.body_state == "formula" else None
    if evidence.definition is None:
        return None
    if evidence.body_state == "values":
        logger.debug(
            "Column '%s' has a calculated-column definition but holds values; treating it as data",
            evidence.name,
        )
        return None
    return evidence.body_formula or "=" + evidence.definition


def openpyxl_formula_columns(
    worksheet: Worksheet,
    table: Table,
    first_body_row: int,
    last_body_row: int,
) -> dict[str, str]:
    """Calculated columns of an openpyxl table, mapped to the formula to write.

    Must run before any body cell is cleared or written.
    """
    definitions = formula_definitions(table)
    first_col = CellRange(table.ref).min_col
    formulas: dict[str, str] = {}
    for offset, column in enumerate(table.tableColumns):
        name = table_column_name(column)
        definition = definitions.get(name)
        if definition is None:
            continue
        kinds = [
            _cell_kind(worksheet._cells.get((row, first_col + offset)))
            for row in range(first_body_row, last_body_row + 1)
        ]
        evidence = ColumnFormulaEvidence(name, definition, True, classify_body(kinds))
        formula = resolve_formula(evidence)
        if formula is not None:
            formulas[name] = formula
    return formulas


def _cell_kind(cell: Any) -> CellKind:  # noqa: ANN401
    if cell is None or cell.value is None:
        return "empty"
    return "formula" if cell.data_type == "f" else "value"


def _is_blank_cell_value(value: Any) -> bool:  # noqa: ANN401
    """Whether *value* should be treated as an empty cell without normalising it."""
    if value is None:
        return True
    if isinstance(value, str):
        return value == ""
    pd = sys.modules.get("pandas")
    if pd is not None and (value is pd.NA or value is pd.NaT):
        return True
    return bool(isinstance(value, float) and math.isnan(value))


def warn_about_ignored_formula_data(
    table_name: str,
    data_columns: Sequence[str],
    rows: Sequence[Sequence[Any]],
    formulas: Mapping[str, str],
) -> list[str]:
    """Warn when the data carries values for calculated columns, which are not written.

    Columns supplied only as blanks (``None``, NaN, ``""``) are ignored silently.
    """
    ignored = [
        column
        for index, column in enumerate(data_columns)
        if column in formulas and any(not _is_blank_cell_value(row[index]) for row in rows)
    ]
    if ignored:
        warn_user(
            f"Table {table_name!r}: calculated column(s) {ignored} keep their formulas; "
            "the data supplied for them was ignored.",
        )
    return ignored
