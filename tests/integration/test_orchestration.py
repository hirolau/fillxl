"""Engine choice, deprecated arguments and how lost template features are reported."""

from __future__ import annotations

import logging
import warnings
from pathlib import Path
from typing import Any

import pytest

from fillxl import CellValueError, FillXLWarning, LossyTemplateWarning, fill
from fillxl._backend import resolve_engine
from tests._assets import CHARTS_TEMPLATE_2
from tests.parity.builders import named_cells_template, table_template

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


def test_win32com_rejects_an_output_format_excel_cannot_save_before_any_work_is_done() -> None:
    """The backend must reject impossible output formats before it launches Excel."""
    with pytest.raises(ValueError, match=r"\.csv"):
        resolve_engine("win32com", Path("t.xlsx"), Path("o.csv"))


def test_an_unknown_engine_is_refused_before_any_file_is_touched(tmp_path: Path) -> None:
    """A typo in the engine name must not leave a half-written output behind."""
    template = _xlsx_template(tmp_path)
    output = tmp_path / "out.xlsx"

    with pytest.raises(ValueError, match="Unknown engine"):
        fill(template, output, engine="openpyxl2", tables={"Sales": _ROWS})  # ty: ignore[invalid-argument-type]  # pyrefly: ignore[bad-argument-type]

    assert not output.exists()


def test_values_and_tables_are_keyword_only(tmp_path: Path) -> None:
    """0.6 deprecated the positional call; 0.7.0 removed it."""
    template = _xlsx_template(tmp_path)

    with pytest.raises(TypeError, match="positional argument"):
        fill(template, tmp_path / "out.xlsx", {}, {"Sales": _ROWS})  # type: ignore[call-arg]  # ty: ignore[too-many-positional-arguments]


def test_values_passed_both_positionally_and_by_keyword_is_an_error(tmp_path: Path) -> None:
    """The call is ambiguous, so Python refuses it as it would for any keyword-only name."""
    template = _xlsx_template(tmp_path)

    with pytest.raises(TypeError, match="positional argument"):
        fill(template, tmp_path / "out.xlsx", {}, values={})  # type: ignore[misc]  # ty: ignore[too-many-positional-arguments]


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


def test_a_format_the_engine_cannot_write_fails_before_the_template_is_loaded(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The error used to arrive at save(), after the whole fill had run."""
    template = _xlsx_template(tmp_path)

    def refuse_load(_self: object, _path: Path) -> None:
        msg = "the template must not be loaded"
        raise AssertionError(msg)

    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: False)
    monkeypatch.setattr("fillxl._backends._openpyxl.OpenpyxlBackend.load", refuse_load)

    with pytest.raises(ValueError, match=r"cannot write a '\.xlsb' file"):
        fill(template, tmp_path / "out.xlsb", tables={"Sales": _ROWS})

    assert not (tmp_path / "out.xlsb").exists()


def test_a_locked_output_says_the_file_is_probably_open_in_excel(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A bare PermissionError from the final rename is the most common Windows error."""
    template = _xlsx_template(tmp_path)

    def locked(_self: Path, _target: Path) -> Path:
        msg = "in use"
        raise PermissionError(msg)

    monkeypatch.setattr(Path, "replace", locked)

    with pytest.raises(PermissionError, match="Is the file open in Excel"):
        fill(template, tmp_path / "out.xlsx", tables={"Sales": _ROWS})


def test_a_named_cell_refuses_tabular_data_and_points_at_tables(tmp_path: Path) -> None:
    """A DataFrame in a named cell used to be written as its text dump."""
    pd = pytest.importorskip("pandas")
    template = named_cells_template(tmp_path / "template.xlsx", cells={"report": ("A1", 0)})

    with pytest.raises(CellValueError, match="tables="):
        fill(template, tmp_path / "out.xlsx", values={"report": pd.DataFrame({"a": [1]})})


def test_a_precision_warning_from_deep_inside_fillxl_points_at_the_caller(
    tmp_path: Path,
) -> None:
    """pytest.warns and -W error match on the caller of fill, not on internals."""
    template = _xlsx_template(tmp_path)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        fill(
            template,
            tmp_path / "out.xlsx",
            tables={"Sales": [{"fund": "A", "nav": 1234567890123456}]},
        )

    precision = [w for w in caught if issubclass(w.category, FillXLWarning)]
    assert len(precision) == 1
    assert Path(precision[0].filename).name == Path(__file__).name


def test_auto_refuses_an_output_suffix_excel_cannot_save(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """``auto`` must not hand Excel a file name it is bound to reject at save()."""
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)

    with pytest.raises(ValueError, match=r"cannot write a '\.csv' file"):
        resolve_engine("auto", Path("t.xlsx"), Path("o.csv"))


def test_auto_never_starts_excel_for_a_format_it_cannot_save(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A bad suffix is refused up front, not after the workbook has been opened.

    Excel can convert between its own formats, not to ``.txt``; the old code
    chose win32com for the format conversion and only found out at ``save()``,
    naming the internal staging file instead of the caller's path.
    """
    template = _xlsx_template(tmp_path)

    def refuse_load(_self: object, _path: Path) -> None:
        msg = "the template must not be loaded"
        raise AssertionError(msg)

    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)
    monkeypatch.setattr("fillxl._backends._win32com.Win32ComBackend.load", refuse_load)
    monkeypatch.setattr("fillxl._backends._openpyxl.OpenpyxlBackend.load", refuse_load)

    with pytest.raises(ValueError, match=r"cannot write a '\.txt' file"):
        fill(template, tmp_path / "out.txt", tables={"Sales": _ROWS})

    assert not (tmp_path / "out.txt").exists()


def test_an_output_without_a_suffix_is_refused_by_name(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The message says "no extension", not the name of a temporary file."""
    template = _xlsx_template(tmp_path)
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)

    with pytest.raises(ValueError, match="without an extension"):
        fill(template, tmp_path / "out", tables={"Sales": _ROWS})


def test_the_message_names_the_engine_the_caller_asked_for(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """``engine="auto"`` must not report a win32com error as if it were requested.

    The caller asked for no particular engine, so the message explains that the
    choice was made for them instead of blaming a backend they never named.
    """
    template = _xlsx_template(tmp_path)
    monkeypatch.setattr("fillxl._backend._win32com_available", lambda: True)

    with pytest.raises(ValueError) as caught:
        fill(template, tmp_path / "out.txt", tables={"Sales": _ROWS})

    message = str(caught.value)
    assert message.startswith("engine='auto' chose win32com")
    assert "'out.txt'" not in message  # the staging file is not the caller's path
    assert "stay on openpyxl" in message


def test_an_output_that_is_a_directory_is_refused_before_the_template_is_opened(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A directory used as ``output=`` is not a permission problem, and is not a fill."""
    template = _xlsx_template(tmp_path)
    target = tmp_path / "out.xlsx"
    target.mkdir()

    def refuse_load(_self: object, _path: Path) -> None:
        msg = "the template must not be loaded"
        raise AssertionError(msg)

    monkeypatch.setattr("fillxl._backends._openpyxl.OpenpyxlBackend.load", refuse_load)

    with pytest.raises(IsADirectoryError, match="is a directory"):
        fill(template, target, tables={"Sales": _ROWS})

    assert list(target.iterdir()) == []


@pytest.mark.parametrize(
    ("argument", "value", "expected"),
    [
        ("values", [("report", "x")], "values="),
        ("tables", [{"fund": "A", "nav": 1}], "tables="),
        ("values", "report", "values="),
        ("config", {"escape_formulas": False}, "config="),
    ],
)
def test_a_mistyped_argument_is_refused_before_the_template_is_opened(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    argument: str,
    value: Any,  # noqa: ANN401
    expected: str,
) -> None:
    """A list, a bare string or a dict of options is a TypeError, not an AttributeError.

    The failure used to surface as ``'list' object has no attribute 'items'``
    from inside the fill, after the template had already been loaded.
    """
    template = _xlsx_template(tmp_path)
    output = tmp_path / "out.xlsx"

    def refuse_load(_self: object, _path: Path) -> None:
        msg = "the template must not be loaded"
        raise AssertionError(msg)

    monkeypatch.setattr("fillxl._backends._openpyxl.OpenpyxlBackend.load", refuse_load)

    with pytest.raises(TypeError, match=expected):
        fill(template, output, **{argument: value})

    assert not output.exists()
