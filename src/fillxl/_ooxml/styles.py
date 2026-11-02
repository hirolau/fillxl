"""Number formats: which style a date needs, and how to add one.

The openpyxl engine gives a date cell without a date format one of the default
formats in :mod:`fillxl._formats`.  ``xl/styles.xml`` is only touched when such
a cell has to appear — a template whose cells already carry the right formats
keeps its style sheet byte for byte.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from openpyxl.styles.numbers import BUILTIN_FORMATS

from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from fillxl._ooxml.package import OoxmlPackage

logger = logging.getLogger(__name__)

STYLES_PART = "xl/styles.xml"
STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"

_CELL_XFS = xmltext.element_re("cellXfs")
_NUM_FMTS = xmltext.element_re("numFmts")
_NUM_FMT = xmltext.element_re("numFmt")
_XF = xmltext.element_re("xf")
_ATTRIBUTE = re.compile(r'\s([\w:.\-]+)="([^"]*)"')

# Attributes that do not describe formatting, so a style can be cloned from a
# template cell without dragging its number format along.
_IGNORED_XF_ATTRIBUTES = ("numFmtId", "applyNumberFormat")


def builtin_format(num_fmt_id: int) -> str | None:
    """The format code of a built-in number format id, if it has one.

    Examples:
        >>> builtin_format(0)
        'General'
        >>> builtin_format(14)
        'mm-dd-yy'
        >>> builtin_format(164) is None
        True

    """
    return BUILTIN_FORMATS.get(num_fmt_id)


def attributes_of(tag: str) -> dict[str, str]:
    """Every attribute of the start tag *tag*, unescaped.

    Examples:
        >>> attributes_of('<xf numFmtId="1" fontId="0"/>')
        {'numFmtId': '1', 'fontId': '0'}

    """
    opening = tag.split(">", 1)[0]
    return {
        match.group(1): xmltext.unescape(match.group(2)) for match in _ATTRIBUTE.finditer(opening)
    }


class StylesPart:
    """Reader/editor for ``xl/styles.xml``."""

    def __init__(self, package: OoxmlPackage) -> None:
        """Bind to the style sheet of *package*; it is loaded on first use."""
        self.package = package
        self._loaded = False
        self._changed = False
        self._text = ""
        self._cell_xfs: list[str] = []
        self._format_codes: dict[int, str] = {}
        self._cache: dict[tuple[int, str], str] = {}

    # ── loading ────────────────────────────────────────────────────────────

    def _load(self) -> None:
        if self._loaded:
            return
        self._loaded = True
        if not self.package.has(STYLES_PART):
            return
        self._text = self.package.text(STYLES_PART)
        block = _CELL_XFS.search(self._text)
        if block is not None:
            self._cell_xfs = [match.group(0) for match in _XF.finditer(block.group(0))]
        num_fmts = _NUM_FMTS.search(self._text)
        if num_fmts is not None:
            for match in _NUM_FMT.finditer(num_fmts.group(0)):
                element = match.group(0)
                identifier = xmltext.attribute(element, "numFmtId")
                code = xmltext.attribute(element, "formatCode")
                if identifier is None or code is None or not identifier.isdigit():
                    continue
                self._format_codes[int(identifier)] = code

    # ── reading ────────────────────────────────────────────────────────────

    def number_format(self, style: str | None) -> str | None:
        """The number-format code a cell style index stands for.

        ``None`` means "unknown": the index is outside the style sheet, or the
        style sheet has no entry for it.  That is what COM returns for a range
        that mixes formats, so callers treat it as "no format information".
        """
        self._load()
        index = int(style) if style is not None and style.isdigit() else 0
        if index >= len(self._cell_xfs):
            return None
        identifier = xmltext.attribute(self._cell_xfs[index], "numFmtId")
        if identifier is None or not identifier.isdigit():
            return "General"
        # A workbook may redefine a built-in id; its own definition wins.
        return self._format_codes.get(int(identifier)) or builtin_format(int(identifier))

    # ── writing ────────────────────────────────────────────────────────────

    def date_style(self, style: str | None, format_code: str) -> str | None:
        """A style index whose number format is *format_code*.

        The style is reused when the template already has one for the same cell
        formatting; otherwise the cell's own style is cloned with the date
        format attached, so borders, fonts and fills survive.

        Args:
            style: Style index the cell has now (``None`` for the default).
            format_code: Excel number-format code, e.g. ``yyyy-mm-dd``.

        Returns:
            The style index to write, or ``None`` when the package has no
            style sheet (the caller then keeps the cell's own style).

        """
        self._load()
        if not self._cell_xfs:
            return None
        source = int(style) if style is not None and style.isdigit() else 0
        key = (source, format_code)
        if key in self._cache:
            return self._cache[key]
        source_xf = self._cell_xfs[source] if source < len(self._cell_xfs) else '<xf numFmtId="0"/>'
        num_fmt_id = self._num_fmt_id(format_code)
        wanted = self._signature(source_xf)
        for index, candidate in enumerate(self._cell_xfs):
            if xmltext.attribute(candidate, "numFmtId") != str(num_fmt_id):
                continue
            if xmltext.attribute(candidate, "applyNumberFormat") == "0":
                continue
            if self._signature(candidate) == wanted:
                self._cache[key] = str(index)
                return str(index)
        clone = xmltext.set_tag_attribute(source_xf, "numFmtId", str(num_fmt_id))
        clone = xmltext.set_tag_attribute(clone, "applyNumberFormat", "1")
        self._cell_xfs.append(clone)
        self._changed = True
        self._cache[key] = str(len(self._cell_xfs) - 1)
        return self._cache[key]

    @staticmethod
    def _signature(xf: str) -> tuple[tuple[str, str], ...]:
        """The attributes of *xf* that describe formatting, minus the number format."""
        return tuple(
            sorted(
                (name, value)
                for name, value in attributes_of(xf).items()
                if name not in _IGNORED_XF_ATTRIBUTES
            ),
        )

    def _num_fmt_id(self, format_code: str) -> int:
        """The ``numFmtId`` for *format_code*, adding a ``numFmt`` when needed."""
        for identifier, code in self._format_codes.items():
            if code == format_code:
                return identifier
        for identifier, code in BUILTIN_FORMATS.items():
            if code == format_code:
                return identifier
        new_id = max([163, *self._format_codes]) + 1
        self._format_codes[new_id] = format_code
        self._changed = True
        element = (
            f'<numFmt numFmtId="{new_id}" formatCode="{xmltext.escape_attribute(format_code)}"/>'
        )
        block = _NUM_FMTS.search(self._text)
        if block is not None:
            content = block.group(0)
            opening = content[: content.find(">") + 1]
            self_closing = opening.endswith("/>")
            if self_closing:
                opening = opening[:-2].rstrip() + ">"
                inner = ""
            else:
                inner = content[content.find(">") + 1 : content.rfind("</")]
            opening = xmltext.set_tag_attribute(opening, "count", str(len(self._format_codes)))
            closing = "</numFmts>"
            rebuilt = f"{opening}{inner}{element}{closing}"
            self._text = self._text[: block.start()] + rebuilt + self._text[block.end() :]
            return new_id
        new_block = f'<numFmts count="1">{element}</numFmts>'
        anchor = xmltext.start_tag_re("fonts").search(self._text)
        position = anchor.start() if anchor is not None else self._text.find(">") + 1
        self._text = f"{self._text[:position]}{new_block}{self._text[position:]}"
        return new_id

    def save(self) -> None:
        """Write ``xl/styles.xml`` back when a date format had to be added."""
        self._load()
        if not self._changed or not self._cell_xfs:
            return
        block = _CELL_XFS.search(self._text)
        if block is None:  # pragma: no cover - guarded by _cell_xfs
            return
        content = block.group(0)
        opening_end = content.find(">")
        closing = content[content.rfind("</") :]
        rebuilt = f"{content[: opening_end + 1]}{''.join(self._cell_xfs)}{closing}"
        rebuilt = xmltext.set_tag_attribute(rebuilt, "count", str(len(self._cell_xfs)))
        self._text = self._text[: block.start()] + rebuilt + self._text[block.end() :]
        self.package.add_override(STYLES_PART, STYLES_CONTENT_TYPE)
        self.package.set_text(STYLES_PART, self._text)
        self._changed = False
