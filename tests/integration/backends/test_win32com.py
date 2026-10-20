"""Win32com CI-safe guard tests (no Excel required).

Tests the no-workbook guard, factory, and platform checks for the
win32com backend.  These do not require Excel and can run in CI.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
import pytest

from fillxl import FillXLConfig


def assets_dir() -> Path:
    """Directory holding the test template assets."""
    return Path(__file__).resolve().parents[2] / "assets"


class TestWin32ComBackendNoWorkbook:
    """Backend method calls without loading a workbook (CI-safe)."""

    @pytest.mark.parametrize(
        "method_call",
        [
            pytest.param(
                lambda b: b.fill_named_cell("test", "value", FillXLConfig()),  # pyrefly: ignore[implicit-any-lambda]
                id="fill_named_cell",
            ),
            pytest.param(
                lambda b: b.fill_table("test", [], FillXLConfig()),  # pyrefly: ignore[implicit-any-lambda]
                id="fill_table",
            ),
            pytest.param(
                lambda b: b.save(Path("output.xlsx")),  # pyrefly: ignore[implicit-any-lambda]
                id="save",
            ),
        ],
    )
    def test_win32com_backend_raises_runtime_error_without_workbook(
        self,
        method_call: object,
    ) -> None:
        """Any backend operation without loading raises RuntimeError."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        with pytest.raises(RuntimeError, match="No workbook loaded"):
            method_call(backend)  # type: ignore[operator]  # ty:ignore[call-non-callable]

    def test_win32com_backend_close_without_load_does_not_raise(self) -> None:
        """Calling close() without loading should not raise."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        backend.close()

    def test_close_without_load_does_not_touch_com(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """close() must not import or call pythoncom when no Excel was started."""
        from fillxl._backends._win32com import Win32ComBackend

        class ForbiddenCom:
            def __getattr__(self, name: str) -> object:
                raise AssertionError(f"pythoncom.{name} was accessed")

        monkeypatch.setitem(sys.modules, "pythoncom", ForbiddenCom())

        Win32ComBackend().close()


class TestWin32ComBackendFactory:
    """Factory tests for the win32com engine (CI-safe)."""

    def test_get_backend_returns_win32com_for_win32com_engine(self) -> None:
        """get_backend('win32com') should return a Win32ComBackend instance."""
        from fillxl._backend import get_backend
        from fillxl._backends._win32com import Win32ComBackend

        backend = get_backend("win32com")
        assert isinstance(backend, Win32ComBackend)

    def test_win32com_backend_does_not_launch_excel_on_init(self) -> None:
        """Win32ComBackend() should not launch Excel until load() is called."""
        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        assert backend._app is None
        assert backend._workbook is None


class TestWin32ComPlatformGuard:
    """Platform guard for non-Windows systems (CI-safe)."""

    def test_win32com_backend_raises_runtime_error_on_non_windows(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """On non-Windows, load() should raise RuntimeError with a clear message."""
        from fillxl._backends import _win32com

        monkeypatch.setattr(_win32com, "_is_windows", False)

        backend = _win32com.Win32ComBackend()
        with pytest.raises(RuntimeError, match="win32com"):
            backend.load(tmp_path / "fake.xlsx")


# ═══════════════════════════════════════════════════════════════════════════════
# Overflow guard (CI-safe, fake COM)
# ═══════════════════════════════════════════════════════════════════════════════


class _FakeRange:
    def __init__(self, value: object, *, fail: bool = False) -> None:
        self._value, self._fail = value, fail

    @property
    def Value(self) -> object:  # noqa: N802
        if self._fail:
            raise OSError("COM read failed")
        return self._value


class _FakeSheet:
    def __init__(self, rng: _FakeRange) -> None:
        self._rng = rng

    def Range(self, address: str) -> _FakeRange:  # noqa: N802
        return self._rng


def _guard(cells: object, *, fail: bool = False) -> None:
    from fillxl._backends._win32com import _check_overflow_com

    _check_overflow_com(
        _FakeSheet(_FakeRange(cells, fail=fail)),
        table_name="T",
        start_col=2,
        start_col_letter="B",
        end_col_letter="C",
        first_row=5,
        last_row=6,
        num_data_rows=4,
    )


def test_overflow_guard_refuses_to_grow_when_cells_cannot_be_read() -> None:
    """An unreadable range must fail closed, not silently overwrite."""
    with pytest.raises(RuntimeError, match="could not be read"):
        _guard(None, fail=True)


def test_overflow_guard_reports_the_first_blocking_cell() -> None:
    """The error names the cell whose content blocks the resize."""
    with pytest.raises(ValueError, match="C6"):
        _guard(((None, None), (None, "x")))


def test_overflow_guard_accepts_an_empty_range() -> None:
    """``None`` and a grid of ``None`` cells both mean "empty"."""
    _guard(None)
    _guard(((None, None), (None, None)))


# ═══════════════════════════════════════════════════════════════════════════════
# Excel-backed behaviours (skipped without Excel)
# ═══════════════════════════════════════════════════════════════════════════════


def _skip_without_excel() -> None:
    """Skip the calling test when this machine cannot drive Excel."""
    if sys.platform != "win32":
        pytest.skip("win32com backend requires Windows")

    try:
        import pythoncom
        import win32com.client

        # A backend fill releases COM on its thread, so direct COM use must init it.
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")
    except Exception as exc:
        pytest.skip(f"Excel COM is not available: {exc}")
    else:
        app.Quit()


class TestWin32ComSaveFormats:
    """The output format follows the requested suffix."""

    def test_filling_an_xlsm_template_keeps_its_macros(self, tmp_path: Path) -> None:
        """An .xlsm template keeps its vbaProject when filled to .xlsm."""
        import zipfile

        from fillxl import fill

        _skip_without_excel()
        template = assets_dir() / "test_template_1_slsm.xlsm"
        output = tmp_path / "macros_kept.xlsm"

        fill(
            template=template,
            output=output,
            values={"single_cell": "Macro Friendly"},
            engine="win32com",
        )

        with zipfile.ZipFile(output) as archive:
            names = archive.namelist()
        assert "xl/vbaProject.bin" in names, names

        wb = openpyxl.load_workbook(output, keep_vba=True)
        try:
            assert wb["TestSingleCell"]["A1"].value == "Macro Friendly"
        finally:
            wb.close()

    def test_filling_to_xlsb_writes_the_binary_format(self, tmp_path: Path) -> None:
        """The .xlsb suffix selects the binary workbook format."""
        import zipfile

        from fillxl import fill

        _skip_without_excel()
        output = tmp_path / "binary.xlsb"

        fill(
            template=assets_dir() / "test_template_1.xlsx",
            output=output,
            values={"single_cell": "Binary"},
            engine="win32com",
        )

        with zipfile.ZipFile(output) as archive:
            names = archive.namelist()
        assert "xl/workbook.bin" in names, names
        assert "xl/workbook.xml" not in names, names

    def test_an_unknown_output_suffix_raises_value_error(self, tmp_path: Path) -> None:
        """A non-Excel suffix is rejected instead of silently writing .xlsx."""
        from fillxl import fill

        _skip_without_excel()
        with pytest.raises(ValueError, match="Supported output formats"):
            fill(
                template=assets_dir() / "test_template_1.xlsx",
                output=tmp_path / "output.csv",
                values={"single_cell": "nope"},
                engine="win32com",
            )


class TestWin32ComCleanup:
    """Excel must not be left running when something fails."""

    def test_excel_is_not_left_running_when_the_template_cannot_be_opened(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A failed load() is still followed by close(), so Excel quits."""
        from fillxl import fill
        from fillxl._backends._win32com import Win32ComBackend

        _skip_without_excel()
        closed: list[bool] = []
        original_close = Win32ComBackend.close

        def _recording_close(self: Win32ComBackend) -> None:
            closed.append(True)
            original_close(self)

        monkeypatch.setattr(Win32ComBackend, "close", _recording_close)

        # A file with a .xlsx suffix that is not a workbook at all.
        broken = tmp_path / "broken.xlsx"
        broken.write_text("this is not a workbook")

        with pytest.raises(Exception):  # noqa: B017
            fill(template=broken, output=tmp_path / "out.xlsx", engine="win32com")

        assert closed == [True], "close() must run even when load() fails"


class TestWin32ComTotalsRows:
    """Tables with totals rows keep the totals row below the data body."""

    @staticmethod
    def _build_template(path: Path, *, data_rows: int) -> None:
        from openpyxl.worksheet.filters import AutoFilter
        from openpyxl.worksheet.table import Table, TableColumn

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        for offset in range(data_rows):
            ws.cell(row=2 + offset, column=1, value=f"old{offset}")
            ws.cell(row=2 + offset, column=2, value=offset)
        totals_row = data_rows + 2
        ws.cell(row=totals_row, column=1, value="Total")
        ws.cell(row=totals_row, column=2, value="=SUBTOTAL(109,Totals[Value])")
        ws.add_table(
            Table(
                displayName="Totals",
                ref=f"A1:B{totals_row}",
                autoFilter=AutoFilter(ref=f"A1:B{totals_row - 1}"),
                tableColumns=[
                    TableColumn(id=1, name="Name", totalsRowLabel="Total"),
                    TableColumn(id=2, name="Value", totalsRowFunction="sum"),
                ],
                totalsRowCount=1,
                totalsRowShown=True,
            ),
        )
        wb.save(path)
        wb.close()

    def test_a_totals_row_stays_at_the_bottom(self, tmp_path: Path) -> None:
        """Resizing keeps the totals row one row below the new data body."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "totals_template.xlsx"
        self._build_template(template, data_rows=1)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={"Totals": [{"Name": f"row{i}", "Value": i} for i in range(1, 4)]},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            table = ws.tables["Totals"]

            assert table.ref == "A1:B5"
            assert table.totalsRowCount == 1
            assert [ws.cell(row=r, column=1).value for r in range(2, 5)] == [
                "row1",
                "row2",
                "row3",
            ]
            assert [ws.cell(row=r, column=2).value for r in range(2, 5)] == [1, 2, 3]
            assert ws.cell(row=5, column=1).value == "Total"
            assert "SUBTOTAL" in str(ws.cell(row=5, column=2).value)
        finally:
            wb.close()

        # Excel recalculates the totals formula over the new body rows.
        wb_values = openpyxl.load_workbook(tmp_path / "out.xlsx", data_only=True)
        try:
            assert wb_values["Data"].cell(row=5, column=2).value == 6
        finally:
            wb_values.close()

    def test_a_totals_row_moves_up_when_the_data_shrinks(self, tmp_path: Path) -> None:
        """Shrinking moves the totals row up without leaving stale rows."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "totals_shrink_template.xlsx"
        self._build_template(template, data_rows=3)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={"Totals": [{"Name": "only", "Value": 7}]},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            table = ws.tables["Totals"]

            assert table.ref == "A1:B3"
            assert ws.cell(row=2, column=1).value == "only"
            assert ws.cell(row=3, column=1).value == "Total"
            assert "SUBTOTAL" in str(ws.cell(row=3, column=2).value)
            assert ws.cell(row=4, column=1).value is None
            assert ws.cell(row=5, column=1).value is None
        finally:
            wb.close()

    def test_growing_by_one_row_keeps_the_totals_row(self, tmp_path: Path) -> None:
        """Growing by a single row does not treat the totals row as overflow."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "totals_grow_one_template.xlsx"
        self._build_template(template, data_rows=1)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={"Totals": [{"Name": "a", "Value": 1}, {"Name": "b", "Value": 2}]},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            assert ws.tables["Totals"].ref == "A1:B4"
            assert [ws.cell(row=r, column=2).value for r in range(2, 4)] == [1, 2]
            assert ws.cell(row=4, column=1).value == "Total"
        finally:
            wb.close()

    def test_content_below_a_totals_table_is_rejected(self, tmp_path: Path) -> None:
        """Content directly below the totals row still blocks table growth."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "totals_overflow_template.xlsx"
        self._build_template(template, data_rows=1)

        wb = openpyxl.load_workbook(template)
        wb["Data"]["A4"] = "DO NOT OVERWRITE"
        wb.save(template)
        wb.close()

        with pytest.raises(ValueError, match="cannot expand"):
            fill(
                template=template,
                output=tmp_path / "out.xlsx",
                tables={"Totals": [{"Name": f"row{i}", "Value": i} for i in range(1, 4)]},
                engine="win32com",
            )


class TestWin32ComHeaderlessTable:
    """Tables whose header row is hidden still fill and keep their ListObject.

    Excel drops the ListObject when a *whole* headerless body is written or
    cleared in one call, so these tests guard the row-by-row workaround.
    """

    @staticmethod
    def _build_template(path: Path, *, data_rows: int) -> None:
        from openpyxl.worksheet.table import Table, TableColumn

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        for offset in range(data_rows):
            ws.cell(row=3 + offset, column=2, value=100 + offset)
            ws.cell(row=3 + offset, column=3, value=offset)
        ws.add_table(
            Table(
                displayName="NoHeaders",
                ref=f"B3:C{2 + data_rows}",
                tableColumns=[
                    TableColumn(id=1, name="Column1"),
                    TableColumn(id=2, name="Column2"),
                ],
                headerRowCount=0,
            ),
        )
        wb.save(path)
        wb.close()

    def test_a_headerless_table_grows_and_keeps_its_list_object(self, tmp_path: Path) -> None:
        """Growing writes every new row and Excel keeps the table."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "headerless_grow.xlsx"
        self._build_template(template, data_rows=2)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={
                "NoHeaders": [
                    {"Column1": 10, "Column2": 60},
                    {"Column1": 20, "Column2": 70},
                    {"Column1": 30, "Column2": 80},
                ],
            },
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            table = ws.tables["NoHeaders"]

            assert table.headerRowCount == 0
            assert table.ref == "B3:C5"
            assert [ws.cell(row=r, column=2).value for r in (3, 4, 5)] == [10, 20, 30]
            assert [ws.cell(row=r, column=3).value for r in (3, 4, 5)] == [60, 70, 80]
        finally:
            wb.close()

    def test_a_headerless_table_can_shrink_to_a_single_row(self, tmp_path: Path) -> None:
        """A one-row body is written cell by cell, so the table survives."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "headerless_shrink.xlsx"
        self._build_template(template, data_rows=3)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={"NoHeaders": [{"Column1": 7, "Column2": 8}]},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            table = ws.tables["NoHeaders"]

            assert table.ref == "B3:C3"
            assert ws.cell(row=3, column=2).value == 7
            assert ws.cell(row=3, column=3).value == 8
            assert ws.cell(row=4, column=2).value is None
        finally:
            wb.close()

    def test_empty_data_keeps_the_headerless_table(self, tmp_path: Path) -> None:
        """Clearing a headerless table row by row does not drop it."""
        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "headerless_empty.xlsx"
        self._build_template(template, data_rows=2)

        fill(
            template=template,
            output=tmp_path / "out.xlsx",
            tables={"NoHeaders": []},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(tmp_path / "out.xlsx")
        try:
            ws = wb["Data"]
            table = ws.tables["NoHeaders"]

            assert table.ref == "B3:C3"
            assert ws.cell(row=3, column=2).value is None
            assert ws.cell(row=3, column=3).value is None
        finally:
            wb.close()


class TestWin32ComOverflowGuard:
    """COM reads the cells below the table before resizing it."""

    def test_a_table_that_would_overwrite_content_is_rejected(
        self,
        tmp_path: Path,
    ) -> None:
        """Expanding over existing cells raises instead of destroying the data."""
        from openpyxl.worksheet.table import Table, TableColumn

        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "overflow_template.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["A2"] = "a"
        ws["B2"] = 1
        # Guard cell directly below the table, in the way of the expansion.
        ws["A5"] = "DO NOT OVERWRITE"
        ws.add_table(
            Table(
                displayName="Blocked",
                ref="A1:B2",
                tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
            ),
        )
        wb.save(template)
        wb.close()

        output = tmp_path / "blocked.xlsx"
        with pytest.raises(ValueError, match=r"cannot expand.*cell A5 "):
            fill(
                template=template,
                output=output,
                tables={"Blocked": [{"Name": f"row{i}", "Value": i} for i in range(1, 10)]},
                engine="win32com",
            )

        assert not output.exists(), "no partial output should be written"

    def test_a_table_with_room_below_it_is_resized(
        self,
        tmp_path: Path,
    ) -> None:
        """The same template without the blocking cell expands normally."""
        from openpyxl.worksheet.table import Table, TableColumn

        from fillxl import fill

        _skip_without_excel()
        template = tmp_path / "roomy_template.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Name"
        ws["B1"] = "Value"
        ws["A2"] = "a"
        ws["B2"] = 1
        ws.add_table(
            Table(
                displayName="Roomy",
                ref="A1:B2",
                tableColumns=[TableColumn(id=1, name="Name"), TableColumn(id=2, name="Value")],
            ),
        )
        wb.save(template)
        wb.close()

        output = tmp_path / "roomy.xlsx"
        fill(
            template=template,
            output=output,
            tables={"Roomy": [{"Name": f"row{i}", "Value": i} for i in range(1, 10)]},
            engine="win32com",
        )

        wb = openpyxl.load_workbook(output)
        try:
            table = next(iter(wb["Data"].tables.values()))
            assert table.ref == "A1:B10", table.ref
            assert wb["Data"]["A10"].value == "row9"
        finally:
            wb.close()
