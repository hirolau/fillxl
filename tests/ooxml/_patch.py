"""Patch one part of a workbook package, keeping every other entry byte for byte.

openpyxl cannot write the shapes that break a surgical editor: shared formulas,
cells without an ``r`` attribute, the 1904 date system, ``extLst`` lookalikes.
The tests therefore start from a workbook openpyxl *can* write and patch the
one part they care about as text.
"""

from __future__ import annotations

import contextlib
import os
import zipfile
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Callable
    from pathlib import Path


def part_text(path: Path, part_name: str) -> str:
    """The decoded text of *part_name* in the package at *path*."""
    with zipfile.ZipFile(path) as archive:
        return archive.read(part_name).decode("utf-8")


def patch_part(path: Path, part_name: str, transform: Callable[[str], str]) -> None:
    """Rewrite one part of the zip at *path* in place, keeping every other entry byte for byte.

    Args:
        path: The workbook to rewrite.
        part_name: The part to transform, e.g. ``xl/worksheets/sheet1.xml``.
        transform: Receives the part's text and returns its replacement.

    Raises:
        KeyError: If the package has no such part.
        AssertionError: If *transform* returned the text unchanged.  A patch
            that silently does nothing makes a test pass for the wrong reason.

    """
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    if part_name not in parts:
        raise KeyError(part_name)
    original = parts[part_name].decode("utf-8")
    updated = transform(original)
    if updated == original:
        msg = f"transform did not change {part_name}"
        raise AssertionError(msg)
    parts[part_name] = updated.encode("utf-8")

    temporary = path.with_name(f"{path.name}.patching")
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for info in infos:
                archive.writestr(info, parts[info.filename])
        os.replace(temporary, path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            temporary.unlink()
