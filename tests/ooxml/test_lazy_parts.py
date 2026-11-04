"""Parts a fill never touches must never be read.

A workbook can hold a 20 MB image; copying the package by reading every part
into memory makes a one-cell fill allocate all of it.  Untouched parts are
streamed from the template into the output instead.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from typing import TYPE_CHECKING

import pytest

from fillxl import fill
from tests.ooxml._build import named_cell_template

if TYPE_CHECKING:
    from pathlib import Path

MEDIA_PREFIX = "xl/media/"


def test_untouched_media_is_never_loaded_into_memory(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The image part is copied without ever being read."""
    template = named_cell_template(tmp_path / "template.xlsx")
    _add_media_part(template, _png(tmp_path / "big.png", 600, 600))
    media = _media_part(template)
    assert _entry_size(template, media) > 500_000

    reads: list[str] = []
    original = zipfile.ZipFile.read

    def counting(archive: zipfile.ZipFile, name: str, pwd: bytes | None = None) -> bytes:
        reads.append(name)
        return original(archive, name, pwd)

    monkeypatch.setattr(zipfile.ZipFile, "read", counting)

    output = fill(template, tmp_path / "out.xlsx", engine="fillxl", values={"report_title": 42})

    assert media not in reads
    assert _entry(output, media) == _entry(template, media)


def _add_media_part(workbook_path: Path, image_path: Path) -> None:
    """Add the image as a package part.

    It is not wired to a drawing: the test is about copying bytes, and a part
    nothing references is never read by either engine.
    """
    with zipfile.ZipFile(workbook_path) as archive:
        infos = archive.infolist()
        parts = {info.filename: archive.read(info.filename) for info in infos}
    media = f"{MEDIA_PREFIX}image1.png"
    temporary = workbook_path.with_name(f"{workbook_path.name}.media")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
        for info in infos:
            archive.writestr(info, parts[info.filename])
        archive.writestr(media, image_path.read_bytes())
    temporary.replace(workbook_path)


def _png(path: Path, width: int, height: int) -> Path:
    """A valid, deliberately poorly compressible PNG."""
    rows = bytearray()
    for y in range(height):
        rows.append(0)
        for x in range(width):
            rows += bytes(((x * 37) % 256, (y * 53) % 256, (x * y) % 256))
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + _chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + _chunk(b"IDAT", zlib.compress(bytes(rows), 0))
        + _chunk(b"IEND", b""),
    )
    return path


def _chunk(tag: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + tag
        + data
        + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    )


def _media_part(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        names = [name for name in archive.namelist() if name.startswith(MEDIA_PREFIX)]
    assert names, "the template has no image"
    return names[0]


def _entry(path: Path, part: str) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read(part)


def _entry_size(path: Path, part: str) -> int:
    with zipfile.ZipFile(path) as archive:
        return archive.getinfo(part).file_size
