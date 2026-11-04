"""``.xls`` (BIFF8) is no longer supported by any engine.

BIFF8 cannot hold the features fillxl exists to preserve, and reading it
needs Excel anyway, so the request is refused before anything is opened.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from fillxl import Engine, fill

if TYPE_CHECKING:
    from pathlib import Path

ENGINES: list[Engine] = ["auto", "openpyxl", "fillxl", "win32com"]


@pytest.mark.parametrize("engine", ENGINES)
def test_an_xls_template_is_refused_before_anything_is_opened(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """The template does not exist, so a different error would mean it was opened."""
    template = tmp_path / "template.xls"

    with pytest.raises(ValueError, match=r"does not support \.xls") as caught:
        fill(template, tmp_path / "out.xlsx", engine=engine)

    assert str(template) in str(caught.value)


@pytest.mark.parametrize("engine", ENGINES)
def test_an_xls_output_is_refused_before_the_template_is_opened(
    engine: Engine,
    tmp_path: Path,
) -> None:
    """Writing a BIFF8 file is as unsupported as reading one."""
    template = tmp_path / "missing.xlsx"

    with pytest.raises(ValueError, match=r"does not support \.xls"):
        fill(template, tmp_path / "out.xls", engine=engine)


def test_an_existing_output_file_is_left_untouched_when_the_request_is_refused(
    tmp_path: Path,
) -> None:
    """A refused request must not truncate the file the caller already has."""
    output = tmp_path / "out.xls"
    output.write_bytes(b"report")

    with pytest.raises(ValueError, match=r"does not support \.xls"):
        fill(tmp_path / "template.xls", output, engine="auto")

    assert output.read_bytes() == b"report"
