"""An OOXML package held in memory, edited part by part.

The package is a ZIP of XML parts.  Reading it into memory lets the engine copy
every part it does not understand byte for byte, which is what keeps spilled
arrays, slicers, rich data and VBA alive: they are never parsed, so they cannot
be lost.
"""

from __future__ import annotations

import posixpath
import re
import shutil
import xml.parsers.expat as expat
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

from attrs import define

from fillxl._ooxml import xmltext

if TYPE_CHECKING:
    from collections.abc import Iterator

CONTENT_TYPES = "[Content_Types].xml"
_ROOT_RELS = "_rels/.rels"
SHARED_STRINGS = "xl/sharedStrings.xml"

_RELATIONSHIP = xmltext.element_re("Relationship")
_SHARED_ITEM = xmltext.element_re("si")
_SHARED_TEXT = xmltext.element_re("t")

_UTF16_LE = b"\xff\xfe"
_UTF16_BE = b"\xfe\xff"

_COPY_CHUNK = 1024 * 1024


def _not_utf8(part: str) -> ValueError:
    """The error raised for a part this engine cannot edit as text."""
    return ValueError(
        f"{part} is not UTF-8 encoded; the fillxl engine only edits UTF-8 parts. "
        "Use engine='win32com'.",
    )


@define(frozen=True)
class Relationship:
    """One entry of a ``.rels`` part.

    Attributes:
        rel_id: The ``Id`` used to reference this relationship.
        rel_type: The full relationship type URI.
        target: The target as written in the part.
        external: ``True`` for ``TargetMode="External"`` targets.

    """

    rel_id: str
    rel_type: str
    target: str
    external: bool = False


def rels_part(part: str) -> str:
    """The ``.rels`` part that lists relationships *part* owns.

    Examples:
        >>> rels_part("xl/workbook.xml")
        'xl/_rels/workbook.xml.rels'
        >>> rels_part("")
        '_rels/.rels'

    """
    if not part:
        return _ROOT_RELS
    folder, _, name = part.rpartition("/")
    return posixpath.join(folder, "_rels", f"{name}.rels")


class OoxmlPackage:
    """A workbook package whose parts can be read, rewritten and dropped."""

    def __init__(self, path: Path) -> None:
        """Read the package at *path*.

        Args:
            path: Path to an ``.xlsx``/``.xlsm`` file.

        Raises:
            ValueError: If the file is not a ZIP package (an encrypted
                workbook, or a binary ``.xls``/``.xlsb`` file).

        """
        if not zipfile.is_zipfile(path):
            msg = (
                f"{path} is not an OOXML (zip) workbook. Password-protected workbooks, "
                ".xls and .xlsb files cannot be filled by the fillxl engine; use "
                "engine='win32com' instead."
            )
            raise ValueError(msg)
        self.path = path
        with zipfile.ZipFile(path) as archive:
            self._order: list[str] = list(archive.namelist())
            self._infos: dict[str, zipfile.ZipInfo] = {
                info.filename: info for info in archive.infolist()
            }
        # Parts are read on first access: a fill touches a handful of XML
        # parts and must not pull a 20 MB image into memory to do it.
        self._parts: dict[str, bytes] = {}
        self._dropped: set[str] = set()
        self._dirty: set[str] = set()
        self._shared_strings: list[str] | None = None

    # ── parts ───────────────────────────────────────────────────────────────

    @property
    def part_names(self) -> list[str]:
        """Every part name, in the order the archive stores them."""
        return [name for name in self._order if self.has(name)]

    def has(self, part: str) -> bool:
        """Whether the package holds *part* (from the archive, or added since)."""
        if part in self._dropped:
            return False
        return part in self._infos or part in self._parts

    def raw(self, part: str) -> bytes:
        """The bytes of *part*, read from the source archive on first access.

        Raises:
            KeyError: If the package has no such part, or it was dropped.

        """
        if part not in self._parts:
            if not self.has(part):
                raise KeyError(part)
            with zipfile.ZipFile(self.path) as archive:
                self._parts[part] = archive.read(part)
        return self._parts[part]

    def text(self, part: str) -> str:
        """*part* decoded as UTF-8 XML.

        Raises:
            KeyError: If the package has no such part.
            ValueError: If the part is not UTF-8 (a UTF-16 byte-order mark, or
                bytes that do not decode).  Editing such a part as text would
                destroy it, and the other engine can read it.

        """
        data = self.raw(part)
        if data.startswith((_UTF16_LE, _UTF16_BE)):
            raise _not_utf8(part)
        try:
            return data.decode("utf-8")
        except UnicodeDecodeError as exc:
            raise _not_utf8(part) from exc

    def set_text(self, part: str, text: str) -> None:
        """Replace *part* with *text*."""
        self._parts[part] = text.encode("utf-8")
        if part not in self._order:
            self._order.append(part)
        self._dropped.discard(part)
        self._dirty.add(part)

    def is_changed(self, part: str) -> bool:
        """Whether *part* was rewritten (as opposed to copied verbatim)."""
        return part in self._dirty

    def shared_strings(self) -> list[str]:
        """The texts of ``xl/sharedStrings.xml``, read on first use.

        Only error messages need them, so the part is not parsed during a
        normal fill.
        """
        if self._shared_strings is None:
            self._shared_strings = self._read_shared_strings()
        return self._shared_strings

    def _read_shared_strings(self) -> list[str]:
        if not self.has(SHARED_STRINGS):
            return []
        text = self.text(SHARED_STRINGS)
        strings: list[str] = []
        for item in _SHARED_ITEM.finditer(text):
            element = item.group(0)
            strings.append(
                "".join(
                    xmltext.element_text(run.group(0), "t") or ""
                    for run in _SHARED_TEXT.finditer(element)
                ),
            )
        return strings

    def drop(self, part: str) -> None:
        """Remove *part*, its relationship part and its content-type override."""
        self._dropped.add(part)
        self._parts.pop(part, None)
        self._order = [name for name in self._order if name != part]
        self._dirty.discard(part)
        rels = rels_part(part)
        if rels != _ROOT_RELS and self.has(rels):
            self._dropped.add(rels)
            self._parts.pop(rels, None)
            self._order = [name for name in self._order if name != rels]
            self._dirty.discard(rels)
        self.remove_override(part)

    # ── relationships ───────────────────────────────────────────────────────

    def relationships(self, owner: str) -> list[Relationship]:
        """The relationships *owner* declares (``""`` means the package root)."""
        part = rels_part(owner)
        if not self.has(part):
            return []
        found: list[Relationship] = []
        text = self.text(part)
        for match in _RELATIONSHIP.finditer(text):
            element = match.group(0)
            rel_id = xmltext.attribute(element, "Id")
            rel_type = xmltext.attribute(element, "Type")
            target = xmltext.attribute(element, "Target")
            if rel_id is None or rel_type is None or target is None:
                continue
            found.append(
                Relationship(
                    rel_id=rel_id,
                    rel_type=rel_type,
                    target=target,
                    external=xmltext.attribute(element, "TargetMode") == "External",
                ),
            )
        return found

    def target_part(self, owner: str, rel_id: str) -> str | None:
        """The package part a relationship of *owner* points at."""
        for relationship in self.relationships(owner):
            if relationship.rel_id != rel_id or relationship.external:
                continue
            return self._resolve(owner, relationship.target)
        return None

    def related_parts(self, owner: str, type_suffix: str) -> list[str]:
        """Every internal part *owner* relates to with a type ending in *type_suffix*."""
        parts: list[str] = []
        for relationship in self.relationships(owner):
            if relationship.external or not relationship.rel_type.endswith(type_suffix):
                continue
            resolved = self._resolve(owner, relationship.target)
            if resolved is not None and self.has(resolved):
                parts.append(resolved)
        return parts

    def remove_relationship(self, owner: str, rel_id: str) -> None:
        """Delete the relationship *rel_id* from *owner*'s ``.rels`` part."""
        part = rels_part(owner)
        if part not in self._parts:
            return
        text = self.text(part)

        def keep(match: re.Match[str]) -> str:
            return "" if xmltext.attribute(match.group(0), "Id") == rel_id else match.group(0)

        self.set_text(part, _RELATIONSHIP.sub(keep, text))

    def dependencies(self, owner: str) -> Iterator[tuple[str, str]]:
        """Yield ``(owner, part)`` for every internal relationship of *owner*."""
        for relationship in self.relationships(owner):
            if relationship.external:
                continue
            resolved = self._resolve(owner, relationship.target)
            if resolved is not None:
                yield owner, resolved

    def reachable_from(self, roots: list[str]) -> set[str]:
        """Every part reachable from *roots* through internal relationships."""
        seen: set[str] = set()
        pending = list(roots)
        while pending:
            part = pending.pop()
            if part in seen or not self.has(part):
                continue
            seen.add(part)
            pending.extend(target for _, target in self.dependencies(part))
        return seen

    @staticmethod
    def _resolve(owner: str, target: str) -> str | None:
        if target.startswith("/"):
            return target.lstrip("/")
        folder = posixpath.dirname(owner)
        resolved = posixpath.normpath(posixpath.join(folder, target))
        return None if resolved.startswith("..") else resolved

    # ── content types ───────────────────────────────────────────────────────

    def remove_override(self, part: str) -> None:
        """Delete the content-type ``<Override>`` of *part*, if there is one."""
        if not self.has(CONTENT_TYPES):
            return
        pattern = re.compile(
            rf'<{xmltext.any_prefix("Override")}\b[^>]*\bPartName="/{re.escape(part)}"[^>]*/>',
        )
        text = self.text(CONTENT_TYPES)
        updated = pattern.sub("", text)
        if updated != text:
            self.set_text(CONTENT_TYPES, updated)

    def add_override(self, part: str, content_type: str) -> None:
        """Declare the content type of *part*."""
        text = self.text(CONTENT_TYPES)
        if f'PartName="/{part}"' in text:
            return
        override = f'<Override PartName="/{part}" ContentType="{content_type}"/>'
        self.set_text(CONTENT_TYPES, text.replace("</Types>", f"{override}</Types>"))

    # ── writing ─────────────────────────────────────────────────────────────

    def save(self, path: Path) -> None:
        """Write the package to *path*.

        Parts that were not rewritten are copied with their original ZIP entry
        (compression method and timestamp included), so an untouched part is
        indistinguishable from the template's.  They are streamed from the
        source archive, so an untouched image is never read into memory.

        The source archive is the temporary copy ``fill()`` makes, and it has
        to still exist here: ``save()`` must stay inside that temporary
        directory's ``with`` block.

        Raises:
            RuntimeError: If a part this engine rewrote is no longer
                well-formed XML.  Nothing is written in that case: half a
                workbook is worse than none.

        """
        self._assert_well_formed()
        # Writing over the source would close the archive being read, so go
        # through a sibling file and swap it in at the end.
        same_file = path.exists() and path.resolve() == self.path.resolve()
        target = path.with_name(f"{path.name}.saving") if same_file else path
        try:
            self._write(target)
        except BaseException:
            if same_file:
                target.unlink(missing_ok=True)
            raise
        if same_file:
            target.replace(path)

    def _write(self, path: Path) -> None:
        """Write every surviving part, streaming the ones that were not touched."""
        with (
            zipfile.ZipFile(self.path) as source,
            zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive,
        ):
            for name in self._order:
                if not self.has(name):
                    continue
                info = self._infos.get(name)
                cached = self._parts.get(name)
                if cached is None and info is not None:
                    # Untouched: copy the entry without materialising it.
                    with source.open(name) as original, archive.open(info, "w") as copy:
                        shutil.copyfileobj(original, copy, length=_COPY_CHUNK)
                    continue
                archive.writestr(info if info is not None else name, cached or b"")

    def _assert_well_formed(self) -> None:
        """Refuse to write a package whose rewritten XML no longer parses.

        The tests assert this about their own output; this guard is what
        protects a user's file, because a text edit gone wrong produces a
        workbook no reader can open.
        """
        for name in sorted(self._dirty):
            if name not in self._parts or not name.endswith((".xml", ".rels")):
                continue
            parser = expat.ParserCreate()
            try:
                parser.Parse(self._parts[name], True)
            except expat.ExpatError as exc:
                msg = (
                    f"fillxl produced malformed XML in {name}: {exc}. "
                    "This is a fillxl bug; please report it with the template."
                )
                raise RuntimeError(msg) from exc
