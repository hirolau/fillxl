"""Engine choice, deprecated arguments and how lost template features are reported."""

from __future__ import annotations

import logging
import warnings
from pathlib import Path

import pytest

from fillxl import LossyTemplateWarning, fill
from fillxl._backend import resolve_engine
from tests._assets import CHARTS_TEMPLATE_2
from tests.parity.builders import table_template

_XLSB_TEMPLATE = Path(__file__).parent.parent / "assets" / "test_template_1_xlsb.xlsb"
_ROWS = [{"fund": "A", "nav": 1}]


def _xlsx_template(tmp_path: Path) -> Path:
    return table_template(tmp_path / "template.xlsx")


def test_auto_keeps_openpyxl_for_an_xlsx_that_stays_an_xlsx(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """There is no reason to start Excel for an .xlsx that stays an .xlsx."""

    def refuse() -> bool:
        msg = "Excel must not be consulted"
        raise AssertionError(msg)

    monkeypatch.setattr("fillxl._backend._win32com_available", refuse)
    template = _xlsx_template(tmp_path)

    fill(template, tmp_path / "out.xlsx", tables={"Sales": _ROWS})

    assert (tmp_path / "out.xlsx").exists()


def test_auto_hands_an_xlsb_template_to_excel_when_excel_is_available(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """openpyxl cannot read .xlsb, so Excel has to do the work."""
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)

    assert resolve_engine("auto", Path("t.xlsb"), Path("o.xlsb")) == "win32com"


def test_auto_lets_excel_convert_between_formats(monkeypatch: pytest.MonkeyPatch) -> None:
    """openpyxl refuses to turn an .xlsx into an .xlsb; Excel can."""
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)

    assert resolve_engine("auto", Path("t.xlsx"), Path("o.xlsb")) == "win32com"


def test_auto_falls_back_to_openpyxl_without_excel_so_the_error_explains_why(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Without Excel the openpyxl error says which engine would have worked."""
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: False)

    with pytest.raises(ValueError, match=r"engine='win32com'"):
        fill(_XLSB_TEMPLATE, tmp_path / "out.xlsb", tables={"TableTest": []})


def test_an_unknown_engine_is_refused_before_any_file_is_touched(tmp_path: Path) -> None:
    """A typo in the engine name must not leave a half-written output behind."""
    template = _xlsx_template(tmp_path)
    output = tmp_path / "out.xlsx"

    with pytest.raises(ValueError, match="Unknown engine"):
        fill(template, output, engine="openpyxl2", tables={"Sales": _ROWS})  # ty: ignore[invalid-argument-type]

    assert not output.exists()


def test_values_and_tables_passed_positionally_still_work_but_warn(tmp_path: Path) -> None:
    """0.5 callers keep working until 0.7.0, with a warning."""
    template = _xlsx_template(tmp_path)
    output = tmp_path / "out.xlsx"

    with pytest.warns(DeprecationWarning, match="positionally"):
        fill(template, output, {}, {"Sales": _ROWS})

    assert output.exists()


def test_values_passed_both_positionally_and_by_keyword_is_an_error(tmp_path: Path) -> None:
    """The call is ambiguous, so it is refused the way Python refuses it."""
    template = _xlsx_template(tmp_path)

    with pytest.raises(TypeError, match="multiple values for argument 'values'"):
        fill(template, tmp_path / "out.xlsx", {}, values={})


def test_lost_template_features_warn_with_a_category_that_can_be_filtered(
    tmp_path: Path,
) -> None:
    """A caller can turn these warnings into errors with one filter."""
    with pytest.warns(LossyTemplateWarning, match="did not preserve"):
        fill(
            CHARTS_TEMPLATE_2,
            tmp_path / "out.xlsx",
            tables={"WaterfallChartTable": [{"Category": "A", "Amount": 100}]},
        )


def test_a_chartex_template_warns_once_not_four_times(
    tmp_path: Path, recwarn: pytest.WarningsRecorder, caplog: pytest.LogCaptureFixture
) -> None:
    """One warning per fill, and no log record saying the same thing again."""
    caplog.set_level(logging.WARNING)

    fill(
        CHARTS_TEMPLATE_2,
        tmp_path / "out.xlsx",
        tables={"WaterfallChartTable": [{"Category": "A", "Amount": 100}]},
    )

    lossy = [warning for warning in recwarn if issubclass(warning.category, LossyTemplateWarning)]
    assert len(lossy) == 1
    assert [record for record in caplog.records if record.levelno >= logging.WARNING] == []


def test_the_lossy_warning_points_at_the_caller_of_fill(tmp_path: Path) -> None:
    """The warning belongs to the person who called fill, not to fillxl internals."""
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        fill(
            CHARTS_TEMPLATE_2,
            tmp_path / "out.xlsx",
            tables={"WaterfallChartTable": [{"Category": "A", "Amount": 100}]},
        )

    lossy = [w for w in caught if issubclass(w.category, LossyTemplateWarning)]
    assert len(lossy) == 1
    assert Path(lossy[0].filename).name == Path(__file__).name


@pytest.mark.excel
def test_the_win32com_engine_never_runs_the_openpyxl_lossy_check(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Excel writes back what it read, so the openpyxl check has nothing to say."""
    calls: list[tuple[object, ...]] = []

    def spy(*args: object, **kwargs: object) -> None:
        calls.append((*args, *kwargs.values()))

    monkeypatch.setattr("fillxl._backends._openpyxl.check_lossy", spy)

    fill(
        CHARTS_TEMPLATE_2,
        tmp_path / "out.xlsx",
        engine="win32com",
        tables={"WaterfallChartTable": [{"Category": "A", "Amount": 100}]},
    )

    assert calls == []
