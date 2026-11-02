"""A lost feature means something different for each engine.

The openpyxl engine is expected to drop what its object model cannot hold, so
the advice is "use win32com".  The fillxl engine copies every part it does not
understand, so losing one is a bug in fillxl, and the message has to say so.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pytest

from fillxl._lossy import LossyTemplateWarning, check_lossy
from tests._spills import build_spill_template
from tests.ooxml._build import table_template

if TYPE_CHECKING:
    from pathlib import Path


def test_a_fillxl_lossy_report_is_described_as_a_bug(tmp_path: Path) -> None:
    """The engine that copies everything cannot lose a part without a bug."""
    template = build_spill_template(tmp_path / "template.xlsx")
    output = table_template(tmp_path / "output.xlsx", rows=1)

    with pytest.warns(LossyTemplateWarning, match="This is a fillxl engine bug"):
        check_lossy(template, output, "warn", engine="fillxl")


def test_the_openpyxl_lossy_wording_is_unchanged(tmp_path: Path) -> None:
    """The openpyxl advice stays exactly as it was."""
    template = build_spill_template(tmp_path / "template.xlsx")
    output = table_template(tmp_path / "output.xlsx", rows=1)

    with pytest.warns(LossyTemplateWarning) as caught:
        check_lossy(template, output, "warn")

    message = str(caught[0].message)
    assert message.startswith("The openpyxl engine did not preserve")
    assert message.endswith("Use engine='win32com' to keep them, or set on_lossy='ignore'.")
