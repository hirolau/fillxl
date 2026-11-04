"""Cell values as ``<c>`` elements.

Everything written here mirrors what the openpyxl engine produces, because the
engines must agree cell for cell: strings become text, ``#N/A`` stays text,
``=1+1`` becomes a formula unless ``escape_formulas`` is on, dates become
serial numbers (with a date number format, applied by
:mod:`fillxl._ooxml.styles`).
"""

from __future__ import annotations

import datetime as dt
from typing import Any

from openpyxl.utils.datetime import WINDOWS_EPOCH, to_excel

from fillxl._normalize import is_text_cell_value
from fillxl._ooxml import xmltext


def serial_text(
    value: dt.date | dt.datetime,
    *,
    epoch: dt.datetime = WINDOWS_EPOCH,
) -> str:
    """The Excel serial number of a date or datetime, as ``<v>`` text.

    Args:
        value: The date or datetime to store.
        epoch: The workbook's day zero.  A workbook saved with the 1904 date
            system counts 1462 days fewer than a 1900-based one.

    Examples:
        >>> serial_text(dt.date(1900, 3, 1))
        '61'
        >>> serial_text(dt.datetime(1900, 3, 1, 12))
        '61.5'
        >>> serial_text(dt.date(1904, 1, 1), epoch=dt.datetime(1904, 1, 1))
        '0'

    """
    serial = to_excel(value, epoch=epoch)
    # Excel stores whole days as integers; writing 61.0 instead is equivalent
    # but noisier in the file.
    return str(int(serial)) if float(serial).is_integer() else repr(serial)


def build_cell(
    reference: str,
    value: Any,  # noqa: ANN401
    *,
    style: str | None,
    escape_formulas: bool,
    epoch: dt.datetime = WINDOWS_EPOCH,
) -> str:
    """The ``<c>`` element holding *value*.

    Args:
        reference: A1 reference such as ``B7``.
        value: An already normalised value (see
            :func:`fillxl._normalize.prepare_cell_value`).
        style: Style index to apply, or ``None`` for the default style.
        escape_formulas: Whether a string starting with ``=`` is text.
        epoch: The workbook's day zero, for date values.

    Examples:
        >>> build_cell("B7", 42, style="1", escape_formulas=False)
        '<c r="B7" s="1"><v>42</v></c>'
        >>> build_cell("B7", 2.5, style=None, escape_formulas=False)
        '<c r="B7"><v>2.5</v></c>'
        >>> build_cell("B7", True, style=None, escape_formulas=False)
        '<c r="B7" t="b"><v>1</v></c>'
        >>> build_cell("B7", "North", style="1", escape_formulas=False)
        '<c r="B7" s="1" t="inlineStr"><is><t xml:space="preserve">North</t></is></c>'
        >>> build_cell("B7", "=A7*2", style="1", escape_formulas=False)
        '<c r="B7" s="1"><f>A7*2</f></c>'
        >>> build_cell("B7", "=A7*2", style="1", escape_formulas=True)
        '<c r="B7" s="1" t="inlineStr"><is><t xml:space="preserve">=A7*2</t></is></c>'

    """
    attributes = f' r="{reference}"'
    if style is not None:
        attributes += f' s="{style}"'
    if value is None:
        return f"<c{attributes}/>"
    if isinstance(value, bool):
        return f'<c{attributes} t="b"><v>{"1" if value else "0"}</v></c>'
    if isinstance(value, (int, float)):
        return f"<c{attributes}><v>{value}</v></c>"
    if isinstance(value, (dt.date, dt.datetime)):
        return f"<c{attributes}><v>{serial_text(value, epoch=epoch)}</v></c>"
    text = str(value)
    if not is_text_cell_value(text, escape_formulas=escape_formulas):
        return f"<c{attributes}><f>{xmltext.escape_text(text[1:])}</f></c>"
    # Excel decodes ``_xHHHH_`` in a string, so literal text has to be escaped
    # the way Excel's own writer escapes it.
    return (
        f'<c{attributes} t="inlineStr"><is>'
        f'<t xml:space="preserve">'
        f"{xmltext.escape_text(xmltext.escape_ooxml_escapes(text))}</t></is></c>"
    )
