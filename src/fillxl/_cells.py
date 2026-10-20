"""Named cell population logic."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from openpyxl.worksheet.worksheet import Worksheet

from fillxl._normalize import check_cell_value, is_text_cell_value, prepare_cell_value

if TYPE_CHECKING:
    from collections.abc import Sequence

    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.workbook.workbook import Workbook

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
        logger.warning("Named cell '%s' has no destinations, skipping", name)
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
        logger.warning(
            "Named cell '%s' has multiple destinations, only the first one (%s!%s) was updated",
            name,
            sheet_name,
            address,
        )
    return sheet_name.replace("''", "'"), address


def find_defined_name(workbook: Workbook, name: str) -> DefinedName | None:
    """Find a defined name, including names scoped to a single sheet.

    ``Workbook.defined_names`` only holds workbook-global names; names scoped
    to a sheet live on that sheet (``Worksheet.defined_names``).  Looking in
    both keeps the openpyxl backend in step with win32com, which resolves
    sheet-scoped names through ``Workbook.Names``.

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

    return None


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

    """
    for name, value in cells.items():
        defined_name = find_defined_name(workbook, name)
        if defined_name is None:
            if allow_missing:
                logger.warning("Named cell '%s' not found in workbook, skipping", name)
                continue
            msg = f"Named cell '{name}' not found in workbook"
            raise KeyError(msg)

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
