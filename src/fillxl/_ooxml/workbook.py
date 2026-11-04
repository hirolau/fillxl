"""The workbook part: sheets, defined names, calculation settings.

Only the attributes and elements fillxl changes are touched; the rest of
``xl/workbook.xml`` is written back exactly as it was read.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from attrs import define

from fillxl._ooxml import xmltext
from fillxl._ooxml.worksheet import CHARTSHEET_VIEW_ANCHORS, root_element_end

if TYPE_CHECKING:
    from fillxl._ooxml.package import OoxmlPackage

logger = logging.getLogger(__name__)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

_SHEETS = xmltext.element_re("sheets")
_SHEET = xmltext.element_re("sheet")
_DEFINED_NAMES = xmltext.element_re("definedNames")
_DEFINED_NAME = xmltext.element_re("definedName")
_CALC_PR = xmltext.element_re("calcPr")
_BOOK_VIEWS = xmltext.element_re("bookViews")
_WORKBOOK_VIEW = xmltext.element_re("workbookView")
_SHEET_VIEWS = xmltext.element_re("sheetViews")
_SHEET_VIEW = xmltext.element_re("sheetView")

_VISIBLE = "visible"
_RELATIONSHIP_SUFFIX = "/officeDocument"

# Element order inside <workbook> (ECMA-376 CT_Workbook): calcPr comes after
# definedNames, so a missing element has to be inserted at the right place.
_AFTER_CALC_PR = (
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "webPublishObjects",
    "extLst",
)


@define(frozen=True)
class SheetEntry:
    """One ``<sheet>`` of the workbook.

    Attributes:
        name: Tab name as Excel shows it.
        sheet_id: The ``sheetId`` attribute.
        rel_id: The relationship id that leads to the sheet part.
        state: ``"visible"``, ``"hidden"`` or ``"veryHidden"``.
        part: Package part holding the sheet (a worksheet or a chartsheet).
        index: Position in ``<sheets>``, which is what ``localSheetId`` uses.

    """

    name: str
    sheet_id: str
    rel_id: str
    state: str
    part: str
    index: int

    @property
    def visible(self) -> bool:
        """Whether Excel shows the tab."""
        return self.state == _VISIBLE


@define(frozen=True)
class DefinedNameEntry:
    """One ``<definedName>`` entry.

    Attributes:
        name: The name as written, including a leading ``_xlnm.`` for the
            built-in names Excel hides.
        text: The unescaped definition, e.g. ``Data!$B$3``.
        local_sheet_id: ``localSheetId`` (an index into ``<sheets>``) when the
            name is scoped to one sheet, otherwise ``None`` for a global name.

    """

    name: str
    text: str
    local_sheet_id: int | None


def _renumber_defined_names(text: str, deleted: list[int]) -> str:
    """*text* with the names of deleted sheets gone and the rest renumbered.

    A name scoped to a deleted sheet has nowhere left to point; every later
    ``localSheetId`` moves up by the number of deleted sheets before it.
    """
    block = _DEFINED_NAMES.search(text)
    if block is None:
        return text

    def keep(match: re.Match[str]) -> str:
        element = match.group(0)
        scope = xmltext.attribute(element, "localSheetId")
        if scope is None or not scope.isdigit():
            return element
        index = int(scope)
        if index in deleted:
            return ""
        shift = sum(1 for gone in deleted if gone < index)
        if not shift:
            return element
        return xmltext.set_tag_attribute(element, "localSheetId", str(index - shift))

    updated = _DEFINED_NAME.sub(keep, block.group(0))
    if _DEFINED_NAME.search(updated) is None:
        return f"{text[: block.start()]}{text[block.end() :]}"
    return f"{text[: block.start()]}{updated}{text[block.end() :]}"


def _reset_first_sheet(text: str, deleted: list[int], remaining: int) -> str:
    """Point ``firstSheet`` back at the first sheet when it no longer exists."""
    view = _WORKBOOK_VIEW.search(text)
    if view is None:  # pragma: no cover - every workbook has a book view
        return text
    element = view.group(0)
    value = xmltext.attribute(element, "firstSheet")
    if value is None or not value.isdigit():
        return text
    index = int(value)
    if index not in deleted and index < remaining:
        return text
    updated = xmltext.set_tag_attribute(element, "firstSheet", "0")
    return f"{text[: view.start()]}{updated}{text[view.end() :]}"


class WorkbookPart:
    """Reader/editor for ``xl/workbook.xml`` and the sheets it lists."""

    def __init__(self, package: OoxmlPackage) -> None:
        """Bind to the workbook part of *package*.

        Args:
            package: The open package.

        Raises:
            ValueError: If the package has no ``officeDocument`` relationship.

        """
        self.package = package
        candidates = [
            relationship
            for relationship in package.relationships("")
            if relationship.rel_type.endswith(_RELATIONSHIP_SUFFIX)
        ]
        if not candidates:
            msg = f"{package.path} is not an Excel workbook: it has no officeDocument part."
            raise ValueError(msg)
        target = package.target_part("", candidates[0].rel_id)
        if target is None or not package.has(target):  # pragma: no cover - defensive
            msg = f"{package.path}: the officeDocument part is missing from the package."
            raise ValueError(msg)
        self.part = target

    # ── raw text ───────────────────────────────────────────────────────────

    @property
    def text(self) -> str:
        """The current XML of the workbook part."""
        return self.package.text(self.part)

    def _set(self, text: str) -> None:
        self.package.set_text(self.part, text)

    # ── sheets ─────────────────────────────────────────────────────────────

    def sheets(self) -> list[SheetEntry]:
        """Every sheet of the workbook, in tab order."""
        block = _SHEETS.search(self.text)
        if block is None:
            return []
        entries: list[SheetEntry] = []
        for index, match in enumerate(_SHEET.finditer(block.group(0))):
            element = match.group(0)
            name = xmltext.attribute(element, "name")
            rel_id = xmltext.attribute(element, "r:id") or xmltext.attribute(element, "id")
            if name is None or rel_id is None:
                continue
            part = self.package.target_part(self.part, rel_id)
            entries.append(
                SheetEntry(
                    name=name,
                    sheet_id=xmltext.attribute(element, "sheetId") or "",
                    rel_id=rel_id,
                    state=xmltext.attribute(element, "state") or _VISIBLE,
                    part=part or "",
                    index=index,
                ),
            )
        return entries

    def sheet(self, name: str) -> SheetEntry | None:
        """The sheet called *name*, matching exactly first and then ignoring case.

        Excel resolves sheet names without regard to case, so a name that
        differs only by case has to find the same sheet here.
        """
        entries = self.sheets()
        for entry in entries:
            if entry.name == name:
                return entry
        folded = name.casefold()
        matches = [entry for entry in entries if entry.name.casefold() == folded]
        return matches[0] if len(matches) == 1 else None

    def worksheet_parts(self) -> list[SheetEntry]:
        """The sheets that are ordinary worksheets (not chart sheets)."""
        return [
            entry
            for entry in self.sheets()
            if entry.part.endswith(".xml") and "/worksheets/" in entry.part
        ]

    def set_sheet_state(self, name: str, state: str) -> None:
        """Give the sheet called *name* the ``state`` attribute *state*."""
        text = self.text
        block = _SHEETS.search(text)
        if block is None:  # pragma: no cover - a workbook always lists sheets
            return

        def replace(match: re.Match[str]) -> str:
            element = match.group(0)
            if xmltext.attribute(element, "name") != name:
                return element
            return xmltext.set_tag_attribute(element, "state", state)

        updated = _SHEET.sub(replace, text[block.start() : block.end()])
        self._set(text[: block.start()] + updated + text[block.end() :])

    # ── defined names ──────────────────────────────────────────────────────

    def defined_names(self) -> list[DefinedNameEntry]:
        """Every global and sheet-scoped defined name, in document order."""
        block = _DEFINED_NAMES.search(self.text)
        if block is None:
            return []
        entries: list[DefinedNameEntry] = []
        for match in _DEFINED_NAME.finditer(block.group(0)):
            element = match.group(0)
            name = xmltext.attribute(element, "name")
            if name is None:
                continue
            scope = xmltext.attribute(element, "localSheetId")
            entries.append(
                DefinedNameEntry(
                    name=name,
                    text=xmltext.element_text(element, "definedName") or "",
                    local_sheet_id=int(scope) if scope is not None and scope.isdigit() else None,
                ),
            )
        return entries

    def set_defined_name_text(
        self,
        name: str,
        local_sheet_id: int | None,
        text: str,
    ) -> None:
        """Give the ``<definedName>`` *name*, in scope *local_sheet_id*, the text *text*.

        A name can exist globally and again on a sheet, so the scope is part of
        what identifies it.
        """
        document = self.text
        block = _DEFINED_NAMES.search(document)
        if block is None:  # pragma: no cover - callers read the name from it
            return

        def replace(match: re.Match[str]) -> str:
            element = match.group(0)
            if xmltext.attribute(element, "name") != name:
                return element
            scope = xmltext.attribute(element, "localSheetId")
            found = int(scope) if scope is not None and scope.isdigit() else None
            if found != local_sheet_id:
                return element
            return xmltext.replace_element_text(element, "definedName", text)

        updated = _DEFINED_NAME.sub(replace, document[block.start() : block.end()])
        self._set(document[: block.start()] + updated + document[block.end() :])

    def find_defined_name(self, name: str) -> DefinedNameEntry | None:
        """Find *name*, matching exactly first and then ignoring case.

        The lookup order matches the openpyxl engine: global names, then names
        scoped to a sheet in tab order, then a case-insensitive match — but
        only when exactly one name matches, because Excel forbids two names
        that differ only by case in the same scope.
        """
        entries = self.defined_names()
        for entry in entries:
            if entry.local_sheet_id is None and entry.name == name:
                return entry
        for entry in entries:
            if entry.local_sheet_id is not None and entry.name == name:
                return entry
        folded = name.casefold()
        scopes: list[int | None] = [None]
        scopes.extend(
            sorted({entry.local_sheet_id for entry in entries if entry.local_sheet_id is not None}),
        )
        for scope in scopes:
            matches = [
                entry
                for entry in entries
                if entry.local_sheet_id == scope and entry.name.casefold() == folded
            ]
            if len(matches) == 1:
                return matches[0]
            if matches:
                return None
        return None

    # ── calculation ────────────────────────────────────────────────────────

    @property
    def date1904(self) -> bool:
        """Whether the workbook counts days from 1904 instead of 1900."""
        tag = xmltext.start_tag(self.text, "workbookPr")
        return tag is not None and xmltext.attribute(tag, "date1904") in ("1", "true")

    def set_full_calc_on_load(self) -> None:
        """Mark the workbook so Excel recalculates every formula when it opens."""
        text = self.text
        existing = _CALC_PR.search(text)
        if existing is not None:
            updated = xmltext.set_tag_attribute(existing.group(0), "fullCalcOnLoad", "1")
            self._set(text[: existing.start()] + updated + text[existing.end() :])
            return
        for name in _AFTER_CALC_PR:
            match = xmltext.start_tag_re(name).search(text)
            if match is not None:
                inserted = f'<calcPr fullCalcOnLoad="1"/>{text[match.start() :]}'
                self._set(text[: match.start()] + inserted)
                return
        closing = text.rfind("</")
        self._set(f'{text[:closing]}<calcPr fullCalcOnLoad="1"/>{text[closing:]}')

    def remove_calc_chain(self) -> None:
        """Delete ``xl/calcChain.xml`` and everything that points at it.

        Excel rebuilds the chain on load; a chain that disagrees with the
        workbook makes it refuse the file instead.
        """
        for relationship in self.package.relationships(self.part):
            if relationship.external:
                continue
            target = self.package.target_part(self.part, relationship.rel_id)
            if target is None or not target.endswith("calcChain.xml"):
                continue
            self.package.remove_relationship(self.part, relationship.rel_id)
            self.package.drop(target)
            logger.debug("Removed %s", target)

    # ── active sheet ───────────────────────────────────────────────────────

    def active_sheet_name(self) -> str | None:
        """The name of the tab Excel shows when the workbook opens."""
        view = _WORKBOOK_VIEW.search(self.text)
        entries = self.sheets()
        if not entries:
            return None
        if view is None:
            return entries[0].name
        index = xmltext.attribute(view.group(0), "activeTab")
        position = int(index) if index is not None and index.isdigit() else 0
        if 0 <= position < len(entries):
            return entries[position].name
        return entries[0].name

    def set_active_sheet(self, name: str) -> None:
        """Point ``activeTab`` at *name*.

        The tab selection lives in each sheet's own part, so the caller sets
        that through the sheet store (worksheets) or
        :meth:`set_tab_selected` (chartsheets).
        """
        entries = self.sheets()
        target = next(entry for entry in entries if entry.name == name)
        self._set_active_tab(target.index)

    def _set_active_tab(self, index: int) -> None:
        """Point ``<workbookView activeTab>`` at *index*."""
        text = self.text
        view = _WORKBOOK_VIEW.search(text)
        if view is not None:
            updated = xmltext.set_tag_attribute(view.group(0), "activeTab", str(index))
            self._set(text[: view.start()] + updated + text[view.end() :])
            return
        new_views = f'<bookViews><workbookView activeTab="{index}"/></bookViews>'
        sheets = _SHEETS.search(text)
        if sheets is not None:
            self._set(f"{text[: sheets.start()]}{new_views}{text[sheets.start() :]}")
            return
        closing = text.rfind("</")  # pragma: no cover - a workbook always lists sheets
        self._set(f"{text[:closing]}{new_views}{text[closing:]}")

    def set_tab_selected(self, part: str, *, selected: bool) -> None:
        """Set ``tabSelected`` on the chartsheet part *part*.

        Chartsheets have no :class:`~fillxl._ooxml.worksheet.SheetPart` (they
        hold no cells), so the workbook edits them directly.
        """
        text = self.package.text(part)
        view = _SHEET_VIEW.search(text)
        value = "1" if selected else "0"
        if view is not None:
            updated = xmltext.set_tag_attribute(view.group(0), "tabSelected", value)
            self.package.set_text(part, text[: view.start()] + updated + text[view.end() :])
            return
        if not selected:
            return
        # A sheet can be saved without a <sheetView>.  <sheetViews> follows
        # <sheetPr> and <dimension> and precedes the row defaults, the columns
        # and the data (on a chartsheet: the page setup and the drawing), so
        # inserting it before the first of those keeps the schema order.
        new_view = '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews>'
        anchors = CHARTSHEET_VIEW_ANCHORS
        for name in anchors:
            anchor = xmltext.start_tag_re(name).search(text)
            if anchor is not None:
                self.package.set_text(
                    part,
                    f"{text[: anchor.start()]}{new_view}{text[anchor.start() :]}",
                )
                return
        views = _SHEET_VIEWS.search(text)
        if views is not None:  # pragma: no cover - a sheetViews block without a view
            self.package.set_text(part, f"{text[: views.start()]}{new_view}{text[views.end() :]}")
            return
        root_end = root_element_end(text)
        self.package.set_text(part, f"{text[:root_end]}{new_view}{text[root_end:]}")

    # ── deleting sheets ────────────────────────────────────────────────────

    def remove_sheets(self, names: list[str]) -> list[SheetEntry]:
        """Delete the sheets called *names* and the parts only they referenced.

        Returns:
            The entries that were removed.

        """
        entries = self.sheets()
        doomed = [entry for entry in entries if entry.name in set(names)]
        if not doomed:
            return []
        # ``localSheetId`` and ``firstSheet`` are positions in <sheets>, so
        # deleting a sheet shifts every later one.
        deleted_indices = sorted(entry.index for entry in doomed)
        remaining = len(entries) - len(doomed)
        orphaned = set().union(
            *(self.package.reachable_from([entry.part]) for entry in doomed if entry.part),
        )
        text = self.text
        for entry in doomed:
            text = self._without_sheet_element(text, entry.name)
            self.package.remove_relationship(self.part, entry.rel_id)
        text = _renumber_defined_names(text, deleted_indices)
        text = _reset_first_sheet(text, deleted_indices, remaining)
        self._set(text)
        for part in sorted(orphaned - self.package.reachable_from([self.part])):
            self.package.drop(part)
        return doomed

    def _without_sheet_element(self, text: str, name: str) -> str:
        """*text* without the ``<sheet>`` element of *name*."""
        block = _SHEETS.search(text)
        if block is None:  # pragma: no cover - a workbook always lists sheets
            return text

        def keep(match: re.Match[str]) -> str:
            return "" if xmltext.attribute(match.group(0), "name") == name else match.group(0)

        updated = _SHEET.sub(keep, text[block.start() : block.end()])
        return text[: block.start()] + updated + text[block.end() :]
