"""Win32com integration tests (require real Excel via COM).

These tests verify that the win32com backend correctly interacts with
the Excel COM interface — lifecycle, named cells, sheet operations,
and table fills.  All tests are marked @pytest.mark.win32com.
"""

from __future__ import annotations

from datetime import datetime, time
from pathlib import Path
from unittest import mock

import openpyxl
import pytest
from openpyxl.worksheet.table import (
    Table,
    TableColumn,
    TableFormula,
)

from fillxl import FillXLConfig, fill


@pytest.mark.win32com
class TestWin32ComLifecycle:
    """COM app lifecycle: load, save, close."""

    def test_win32com_backend_loads_workbook_into_excel(self, output_path: Path) -> None:
        """Loading a workbook should open it in Excel."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            assert backend._workbook is not None
        finally:
            backend.close()

    def test_win32com_backend_saves_workbook_to_output_file(self, output_path: Path) -> None:
        """Saving a workbook should create the output file."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            backend.save(output_path)
            assert output_path.exists()
            wb = openpyxl.load_workbook(output_path)
            wb.close()
        finally:
            backend.close()

    def test_win32com_backend_close_releases_excel_process(self) -> None:
        """Closing the backend should release the Excel process."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        backend.load(TEMPLATE)
        backend.close()
        assert backend._app is None
        assert backend._workbook is None


@pytest.mark.win32com
class TestWin32ComNamedCells:
    """Named cell operations via win32com."""

    def test_named_cells_are_filled_via_excel_com(self, output_path: Path) -> None:
        """Fill a named cell and verify via openpyxl read-back."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            backend.fill_named_cell("single_cell", "COM Value", FillXLConfig())
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestSingleCell"]
                assert ws["A1"].value == "COM Value"
            finally:
                wb.close()
        finally:
            backend.close()

    def test_keyerror_raised_for_nonexistent_named_cell(self) -> None:
        """Filling a nonexistent named cell should raise KeyError."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            with pytest.raises(KeyError, match="Nonexistent"):
                backend.fill_named_cell("NonexistentCell", "value", FillXLConfig())
        finally:
            backend.close()

    def test_multiple_named_cells_filled_in_one_session(self, output_path: Path) -> None:
        """Filling multiple named cells should update all of them."""
        from datetime import datetime

        from conftest import _build_multi_named_cells_workbook

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "multi_cells_com.xlsx"
        _build_multi_named_cells_workbook(template)

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_named_cell("single_cell", "Updated", FillXLConfig())
            backend.fill_named_cell("date_cell", datetime(2025, 6, 15), FillXLConfig())
            backend.fill_named_cell("bool_cell", True, FillXLConfig())
            backend.fill_named_cell("pct_cell", 0.42, FillXLConfig())
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestSingleCell"]
                assert ws["A1"].value == "Updated"
                val = ws["B1"].value
                if hasattr(val, "year"):
                    assert val.year == 2025
                    assert val.month == 6
                    assert abs(val.day - 15) <= 1
                else:
                    assert val == datetime(2025, 6, 15)
                assert ws["C1"].value is True
                assert ws["D1"].value == 0.42
            finally:
                wb.close()
        finally:
            backend.close()


@pytest.mark.win32com
class TestWin32ComSheetOps:
    """Sheet operations via win32com."""

    def test_underscore_sheets_hidden_via_excel_com(self, output_path: Path) -> None:
        """Hide sheets starting with '_'."""
        from conftest import _build_hidden_sheets_workbook

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "hidden_com.xlsx"
        _build_hidden_sheets_workbook(template)

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.hide_sheets("_")
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                assert wb["_HiddenSheet"].sheet_state == "hidden"
                assert wb["NormalSheet"].sheet_state == "visible"
            finally:
                wb.close()
        finally:
            backend.close()

    def test_sheets_deleted_via_excel_com(self, output_path: Path) -> None:
        """Delete specified sheets."""
        from conftest import _build_delete_sheets_workbook

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "delete_com.xlsx"
        _build_delete_sheets_workbook(template)

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.delete_sheets(["ToDelete"])
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                assert "ToDelete" not in wb.sheetnames
                assert "TestSheet" in wb.sheetnames
            finally:
                wb.close()
        finally:
            backend.close()

    def test_active_sheet_set_via_excel_com(self, output_path: Path) -> None:
        """Set the active worksheet."""
        from conftest import _build_hidden_sheets_workbook

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "active_com.xlsx"
        _build_hidden_sheets_workbook(template)

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.set_active_sheet("NormalSheet")
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                assert wb.active.title == "NormalSheet"
            finally:
                wb.close()
        finally:
            backend.close()


@pytest.mark.win32com
class TestWin32ComTableFill:
    """Table fill operations via win32com."""

    @staticmethod
    def _data_rows(count: int) -> list[dict[str, object]]:
        return [
            {
                "Col1": f"row{i}",
                "Col2": i,
                "ColNotFilled": f"nf{i}",
                "ColFormula": None,
                "TestFormattedPct": 0.0,
            }
            for i in range(1, count + 1)
        ]

    def test_tables_grow_when_data_is_larger_than_excel_table(self, output_path: Path) -> None:
        """Table should grow when data has more rows than the original."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            spec = backend.fill_table("TableTest", self._data_rows(10), FillXLConfig())
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestTable"]
                assert ws.cell(row=2, column=1).value == "row1"
                assert ws.cell(row=11, column=1).value == "row10"
                assert ws.cell(row=2, column=2).value == 1
                assert ws.cell(row=11, column=2).value == 10
                table = ws.tables["TableTest"]
                assert table.ref == "A1:E11"
            finally:
                wb.close()

            assert spec.old_ref != spec.new_ref
            assert spec.worksheet_title == "TestTable"
        finally:
            backend.close()

    def test_tables_shrink_when_data_has_fewer_rows_than_excel_table(
        self,
        output_path: Path,
    ) -> None:
        """Table should shrink when data has fewer rows than the original."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            spec = backend.fill_table("TableTest", self._data_rows(1), FillXLConfig())
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestTable"]
                assert ws.cell(row=2, column=1).value == "row1"
                assert (
                    ws.cell(row=3, column=1).value is None or ws.cell(row=3, column=1).value == ""
                )
                table = ws.tables["TableTest"]
                assert table.ref == "A1:E2"
            finally:
                wb.close()

            assert spec.new_ref is not None
        finally:
            backend.close()

    def test_table_styles_copied_to_new_rows_when_enabled(self, output_path: Path) -> None:
        """Table fill should copy styles to new rows when enabled."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            backend.fill_table(
                "TableTest",
                self._data_rows(5),
                FillXLConfig(copy_table_styles=True),
            )
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestTable"]
                assert ws.cell(row=6, column=1).value == "row5"
                assert ws.cell(row=6, column=2).value == 5
            finally:
                wb.close()
        finally:
            backend.close()

    def test_calculated_column_formulas_propagate_to_new_rows(self, output_path: Path) -> None:
        """Calculated column formulas should propagate to new rows."""
        from conftest import _build_calculated_column_workbook

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "calc_com.xlsx"
        _build_calculated_column_workbook(template)

        data = [
            {"Col1": "Alice", "Col2": 10, "ColNotFilled": "x", "TestFormattedPct": 0.0},
            {"Col1": "Bob", "Col2": 20, "ColNotFilled": "y", "TestFormattedPct": 0.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table("TableTest", data, FillXLConfig())
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestTable"]
                for row in range(2, 4):
                    val = ws.cell(row=row, column=4).value
                    assert val is not None, f"Row {row} ColFormula is empty"
                    assert isinstance(val, str) and val.startswith("="), (
                        f"Row {row} ColFormula is not a formula: {val!r}"
                    )
            finally:
                wb.close()
        finally:
            backend.close()


def _com_column_has_formula(path: Path, table: str, column: str) -> object:
    """Open *path* in a separate Excel and return ``HasFormula`` of a table column."""
    import pythoncom
    import win32com.client

    # A backend fill releases COM on its thread, so direct COM use must init it.
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.AutomationSecurity = 3
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
        try:
            for sheet in wb.Worksheets:
                for list_obj in sheet.ListObjects:
                    if list_obj.Name == table:
                        return list_obj.ListColumns(column).DataBodyRange.HasFormula
        finally:
            wb.Close(SaveChanges=False)
    finally:
        app.Quit()
    raise KeyError(table)


def _table_rows(count: int) -> list[dict[str, object]]:
    """Rows for the five-column ``TableTest`` table, without ``ColFormula``."""
    return [
        {
            "Col1": f"row{i}",
            "Col2": i,
            "ColNotFilled": f"nf{i}",
            "TestFormattedPct": 0.0,
        }
        for i in range(1, count + 1)
    ]


@pytest.mark.win32com
class TestWin32ComNamedCellParity:
    """Named-cell parity with the openpyxl backend."""

    def test_writing_to_a_protected_sheet_raises_runtime_error_not_keyerror(
        self,
        tmp_path: Path,
    ) -> None:
        """A write failure is not reported as a missing name.

        openpyxl does not enforce sheet protection; this divergence is
        intentional and documented.
        """
        from openpyxl.workbook.defined_name import DefinedName

        template = tmp_path / "protected.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Guarded"
        ws["A1"] = "before"
        wb.defined_names.add(DefinedName("guarded", attr_text="Guarded!$A$1"))
        ws.protection.sheet = True
        wb.save(template)
        wb.close()

        with pytest.raises(RuntimeError, match="guarded"):
            fill(
                template=template,
                output=tmp_path / "protected_out.xlsx",
                values={"guarded": "after"},
                engine="win32com",
            )

    def test_a_missing_name_still_raises_keyerror(self, output_path: Path) -> None:
        """``KeyError`` means "no such name", nothing else."""
        from conftest import TEMPLATE

        with pytest.raises(KeyError, match="does_not_exist"):
            fill(
                template=TEMPLATE,
                output=output_path,
                values={"does_not_exist": 1},
                engine="win32com",
            )


@pytest.mark.win32com
class TestWin32ComSecurityAndThreading:
    """Excel starts hardened and can be driven from worker threads."""

    def test_excel_starts_with_macros_events_and_link_updates_disabled(self) -> None:
        """The security settings are applied before any workbook opens."""
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            assert backend._app.AutomationSecurity == 3
            assert backend._app.EnableEvents is False
            assert backend._app.AskToUpdateLinks is False
        finally:
            backend.close()

    def test_fill_works_from_a_worker_thread(self, output_path: Path) -> None:
        """fill() initialises COM on the thread that uses it."""
        from concurrent.futures import ThreadPoolExecutor

        from conftest import TEMPLATE

        with ThreadPoolExecutor(max_workers=1) as pool:
            result = pool.submit(
                fill,
                template=TEMPLATE,
                output=output_path,
                values={"single_cell": "x"},
                engine="win32com",
            ).result()

        assert result.exists()

    def test_two_sequential_fills_on_the_same_worker_thread(self, tmp_path: Path) -> None:
        """CoInitialize/CoUninitialize stay balanced across fills."""
        from conftest import TEMPLATE

        def run(number: int) -> Path:
            output = tmp_path / f"threaded_{number}.xlsx"
            fill(
                template=TEMPLATE,
                output=output,
                values={"single_cell": f"x{number}"},
                engine="win32com",
            )
            return output

        from concurrent.futures import ThreadPoolExecutor

        with ThreadPoolExecutor(max_workers=1) as pool:
            results = [pool.submit(run, number).result() for number in range(2)]

        assert all(path.exists() for path in results)


@pytest.mark.win32com
def test_a_filled_time_works_in_excel_arithmetic(tmp_path: Path) -> None:
    """A time is a fraction of a day, so ``=A1*24`` returns the hour."""
    from openpyxl.workbook.defined_name import DefinedName

    template = tmp_path / "time_arithmetic.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = None
    ws["A1"].number_format = "hh:mm:ss"
    ws["B1"] = "=A1*24"
    wb.defined_names.add(DefinedName("t", attr_text=f"{ws.title}!$A$1"))
    wb.save(template)
    wb.close()

    output = tmp_path / "time_arithmetic_filled.xlsx"
    fill(
        template=template,
        output=output,
        engine="win32com",
        values={"t": time(9, 15, 30)},
    )

    result = openpyxl.load_workbook(output, data_only=True)
    try:
        assert result[ws.title]["B1"].value == pytest.approx(9.258333, abs=1e-6)
    finally:
        result.close()


@pytest.mark.win32com
class TestWin32ComCalculatedColumns:
    """Calculated columns keep their formulas on every template format."""

    @staticmethod
    def _assets() -> Path:
        return Path(__file__).resolve().parent.parent / "assets"

    def test_xlsb_template_keeps_its_calculated_column(self, tmp_path: Path) -> None:
        """Without readable definitions, a fully formula body still counts."""
        output = tmp_path / "filled.xlsb"
        fill(
            template=self._assets() / "test_template_1_xlsb.xlsb",
            output=output,
            engine="win32com",
            tables={"TableTest": _table_rows(3)},
        )

        assert _com_column_has_formula(output, "TableTest", "ColFormula") is True

    def test_xlsb_drop_strategy_keeps_its_calculated_column(self, tmp_path: Path) -> None:
        """The drop strategy must not turn the formula column into data."""
        output = tmp_path / "filled_drop.xlsb"
        fill(
            template=self._assets() / "test_template_1_xlsb.xlsb",
            output=output,
            engine="win32com",
            tables={"TableTest": _table_rows(3)},
            config=FillXLConfig(column_validation_strategy="drop"),
        )

        assert _com_column_has_formula(output, "TableTest", "ColFormula") is True

    def test_data_for_a_formula_column_is_never_written_as_none(self, tmp_path: Path) -> None:
        """Every body cell of a calculated column stays a formula."""
        output = tmp_path / "filled_formula.xlsx"
        rows = [dict(row, ColFormula=None) for row in _table_rows(5)]
        fill(
            template=self._assets() / "test_template_1.xlsx",
            output=output,
            engine="win32com",
            tables={"TableTest": rows},
        )

        assert _com_column_has_formula(output, "TableTest", "ColFormula") is True


@pytest.mark.win32com
class TestWin32ComFormulaFixes:
    """Regression tests for formula-related bugs in the win32com backend.

    These tests verify that fixes for formula handling are not accidentally
    reverted.
    """

    # ── Helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _build_workbook(
        path: Path,
        *,
        formula_attr: str = 'Col1&" "&Col2',
    ) -> None:
        """Build a workbook with a ``TableTest`` table that has one
        calculated column.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TestTable"

        headers = ["Col1", "Col2", "ColNotFilled", "ColFormula", "TestFormattedPct"]
        col_defs = [
            TableColumn(id=1, name="Col1"),
            TableColumn(id=2, name="Col2"),
            TableColumn(id=3, name="ColNotFilled"),
            TableColumn(
                id=4,
                name="ColFormula",
                calculatedColumnFormula=TableFormula(attr_text=formula_attr),
            ),
            TableColumn(id=5, name="TestFormattedPct"),
        ]

        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        for col_idx in range(1, 6):
            ws.cell(row=2, column=col_idx, value=0)
        # The body must hold a formula for the column to count as calculated.
        ws.cell(row=2, column=4, value=f"={formula_attr}")

        table = Table(
            displayName="TableTest",
            ref="A1:E2",
            tableColumns=col_defs,
        )
        ws.add_table(table)
        wb.save(path)
        wb.close()

    # ── Tests ────────────────────────────────────────────────────────────

    def test_formula_column_formatting_preserved_via_win32com(
        self,
        output_path: Path,
    ) -> None:
        """Multi-line calculated formulas keep newlines and indentation.

        Users often format long ``LET`` / ``LAMBDA`` formulas with line
        breaks for readability.  The win32com backend must not flatten
        them when re-applying formulas after a table fill.
        """
        from fillxl._backends._win32com import Win32ComBackend

        multi_line = (
            'LET(\n        left, Col1,\n        right, Col2,\n        left&" "&right\n    )'
        )
        template = output_path.parent / "formula_formatted.xlsx"
        self._build_workbook(template, formula_attr=multi_line)

        data = [
            {"Col1": "A", "Col2": 10, "ColNotFilled": "x", "TestFormattedPct": 0.0},
            {"Col1": "B", "Col2": 20, "ColNotFilled": "y", "TestFormattedPct": 0.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table("TableTest", data, FillXLConfig())
            backend.save(output_path)
        finally:
            backend.close()

        wb = openpyxl.load_workbook(output_path)
        try:
            ws = wb["TestTable"]
            formula = ws.cell(row=2, column=4).value
            assert isinstance(formula, str)
            assert formula.startswith("=")
            assert "\n" in formula, f"newlines lost: {formula!r}"
            # Excel may rewrite names (e.g. Col1→COL1, add _xlpm.), but the
            # original multi-line indentation must survive.
            assert "\n        " in formula, f"indentation lost: {formula!r}"
            assert formula.count("\n") >= 3
        finally:
            wb.close()

    def test_formula_with_xlfn_prefix_does_not_crash(self, output_path: Path) -> None:
        """A formula containing ``_xlfn.`` / ``_xlpm.`` prefixes must not crash.

        These prefixes are openpyxl XML artifacts that Excel COM does not
        understand and must be stripped before assignment.
        """
        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "xlfn_formula.xlsx"
        self._build_workbook(
            template,
            formula_attr="_xlfn.LET(_xlpm.x, 1, _xlpm.x+1)",
        )

        data = [
            {"Col1": "A", "Col2": 10, "ColNotFilled": "x", "TestFormattedPct": 0.0},
            {"Col1": "B", "Col2": 20, "ColNotFilled": "y", "TestFormattedPct": 0.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            # Must not raise
            backend.fill_table("TableTest", data, FillXLConfig())
            backend.save(output_path)
        finally:
            backend.close()

    def test_plain_number_default_not_treated_as_formula(self, output_path: Path) -> None:
        """A column with ``calculatedColumnFormula='100'`` must not be
        treated as a formula column — the data value must be written
        instead of ``=100``.
        """
        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "plain_default.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "PriceSheet"
        for col_idx, h in enumerate(["Date", "StockPrice"], 1):
            ws.cell(row=1, column=col_idx, value=h)
        ws.cell(row=2, column=1, value=datetime(2024, 1, 1))
        ws.cell(row=2, column=2, value=0)

        table = openpyxl.worksheet.table.Table(
            displayName="PriceTable",
            ref="A1:B2",
            tableColumns=[
                openpyxl.worksheet.table.TableColumn(id=1, name="Date"),
                openpyxl.worksheet.table.TableColumn(
                    id=2,
                    name="StockPrice",
                    calculatedColumnFormula=openpyxl.worksheet.table.TableFormula(attr_text="100"),
                ),
            ],
        )
        ws.add_table(table)
        wb.save(template)
        wb.close()

        data = [
            {"Date": datetime(2025, 1, 1), "StockPrice": 42.5},
            {"Date": datetime(2025, 2, 1), "StockPrice": 55.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table("PriceTable", data, FillXLConfig())
            backend.save(output_path)

            wb2 = openpyxl.load_workbook(output_path)
            try:
                ws2 = wb2["PriceSheet"]
                val = ws2["B3"].value
                assert isinstance(val, (int, float)), f"Expected number, got {val!r}"
                assert abs(float(val) - 55.0) < 0.01, (
                    "StockPrice should be 55.0, not overwritten by =100"
                )
            finally:
                wb2.close()
        finally:
            backend.close()

    def test_datetime_timezone_preserved_via_win32com(self, output_path: Path) -> None:
        """A naive datetime written via win32com must read back with the
        same year/month/day (no UTC+1 timezone shift).
        """
        from conftest import _build_cross_backend_dtype_template

        from fillxl._backends._win32com import Win32ComBackend

        template = output_path.parent / "dtype_com.xlsx"
        _build_cross_backend_dtype_template(template)

        data = [{"DatetimeCol": datetime(2025, 6, 15, 14, 30, 0)}]
        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table(
                "CrossDtypeTable",
                data,
                FillXLConfig(column_validation_strategy="none"),
            )
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["CrossDtypeTable"]
                val = ws["F2"].value
                assert val is not None, "DatetimeCol is empty"
                assert hasattr(val, "year"), f"Expected datetime, got {type(val)}"
                assert val.year == 2025, f"Year mismatch: {val.year}"
                assert val.month == 6, f"Month mismatch: {val.month}"
                assert val.day == 15, f"Day mismatch: {val.day}"
            finally:
                wb.close()
        finally:
            backend.close()


class TestReadTableFormulaDefinitions:
    """Unit tests for ``_read_table_formula_definitions``.

    These tests read workbooks with openpyxl (no Excel COM required).  They
    verify that the helper returns the raw definitions of every table, and
    that formats it cannot read report ``None`` instead of guessing.
    """

    def _make_table_wb(
        self,
        path: Path,
        col_formulas: dict[str, str | None],
    ) -> None:
        """Build a one-row workbook with a ``TestTable`` containing the
        given columns and their ``calculatedColumnFormula`` values.

        Pass ``None`` for columns without a formula.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        col_defs: list[openpyxl.worksheet.table.TableColumn] = []
        for idx, (name, formula) in enumerate(col_formulas.items(), 1):
            ws.cell(row=1, column=idx, value=name)
            ws.cell(row=2, column=idx, value=0)
            if formula is not None:
                col_defs.append(
                    openpyxl.worksheet.table.TableColumn(
                        id=idx, name=name, calculatedColumnFormula=TableFormula(attr_text=formula)
                    )
                )
            else:
                col_defs.append(openpyxl.worksheet.table.TableColumn(id=idx, name=name))

        table = openpyxl.worksheet.table.Table(
            displayName="TestTable",
            ref=f"A1:{openpyxl.utils.get_column_letter(len(col_formulas))}2",
            tableColumns=col_defs,
        )
        ws.add_table(table)
        wb.save(path)
        wb.close()

    # ── Tests ────────────────────────────────────────────────────────────

    def test_xlsx_definitions_are_read(self, tmp_path: Path) -> None:
        r"""A column with ``calculatedColumnFormula='Col1&\" \"&Col2'`` is kept."""
        from fillxl._backends._win32com import _read_table_formula_definitions

        p = tmp_path / "real_formula.xlsx"
        self._make_table_wb(p, {"Col1": None, "ColFormula": 'Col1&" "&Col2'})

        assert _read_table_formula_definitions(p) == {
            "TestTable": {"ColFormula": 'Col1&" "&Col2'},
        }

    def test_xlsb_paths_have_no_definitions(self, tmp_path: Path) -> None:
        """``.xlsb`` hides definitions; callers must fall back to cell formulas."""
        from fillxl._backends._win32com import _read_table_formula_definitions

        p = tmp_path / "book.xlsb"
        p.write_bytes(b"not a workbook")

        assert _read_table_formula_definitions(p) is None

    def test_a_missing_file_has_no_definitions(self, tmp_path: Path) -> None:
        """A missing path is reported as unreadable, not as "no formulas"."""
        from fillxl._backends._win32com import _read_table_formula_definitions

        assert _read_table_formula_definitions(tmp_path / "missing.xlsx") is None

    def test_a_corrupt_xlsx_returns_none_with_a_warning(
        self,
        tmp_path: Path,
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """A corrupt workbook is unreadable and says so."""
        import logging

        from fillxl._backends._win32com import _read_table_formula_definitions

        p = tmp_path / "corrupt.xlsx"
        p.write_bytes(b"not a zip file")

        with caplog.at_level(logging.WARNING, logger="fillxl._backends._win32com"):
            assert _read_table_formula_definitions(p) is None
        assert "Could not read table definitions" in caplog.text

    def test_all_tables_are_read_with_one_workbook_load(
        self,
        tmp_path: Path,
    ) -> None:
        """Filling many tables must not reload the workbook once per table."""
        from fillxl._backends._win32com import Win32ComBackend

        p = tmp_path / "two_tables.xlsx"
        wb = openpyxl.Workbook()
        table_defs = (
            ("First", "TableA", "Col1+1"),
            ("Second", "TableB", 'Col1&"!"'),
        )
        for sheet_name, table_name, formula in table_defs:
            ws = wb.create_sheet(sheet_name)
            ws["A1"] = "Col1"
            ws["A2"] = 0
            ws.add_table(
                Table(
                    displayName=table_name,
                    ref="A1:A2",
                    tableColumns=[
                        TableColumn(
                            id=1,
                            name="Col1",
                            calculatedColumnFormula=TableFormula(attr_text=formula),
                        ),
                    ],
                ),
            )
        wb.save(p)
        wb.close()

        backend = Win32ComBackend()
        backend._load_path = p

        with mock.patch.object(
            openpyxl,
            "load_workbook",
            wraps=openpyxl.load_workbook,
        ) as load_workbook:
            assert backend._table_formula_definitions("TableA") == {"Col1": "Col1+1"}
            assert backend._table_formula_definitions("TableB") == {"Col1": 'Col1&"!"'}

        assert load_workbook.call_count == 1

    def test_prepare_formula_for_com_preserves_newlines_and_indent(
        self,
    ) -> None:
        """``_prepare_formula_for_com`` keeps LF indentation and strips
        openpyxl XML prefixes only.
        """
        from fillxl._backends._win32com import _prepare_formula_for_com

        raw = (
            "=_xlfn.LET(\n"
            "        _xlpm.dates, StockPriceHistory[Date],\n"
            "        _xlpm.x, 1,\n"
            "        _xlpm.x+1\n"
            "    )"
        )
        prepared = _prepare_formula_for_com(raw)
        assert prepared.startswith("=LET(\n")
        assert "\n        dates, StockPriceHistory[Date]," in prepared
        assert "_xlfn." not in prepared
        assert "_xlpm." not in prepared
        assert prepared.count("\n") == raw.count("\n")

    def test_prepare_formula_for_com_normalizes_crlf(self) -> None:
        """CRLF / bare CR are normalized to LF (Excel COM rejects CR)."""
        from fillxl._backends._win32com import _prepare_formula_for_com

        assert _prepare_formula_for_com("=A1\r\n+B1") == "=A1\n+B1"
        assert _prepare_formula_for_com("=A1\r+B1") == "=A1\n+B1"


@pytest.mark.win32com
class TestWin32ComFormatPreservation:
    """Format-preservation fixes for the win32com backend.

    These tests verify that NumberFormat is correctly preserved per column
    after table fills, and that the ``copy_table_styles`` and conditional
    formatting extension steps don't bleed date formatting into numeric
    columns.
    """

    def test_per_column_number_format_preserved_after_table_fill(
        self,
        output_path: Path,
    ) -> None:
        """After filling a table with mixed types, each column must retain
        its original NumberFormat.
        """
        from fillxl._backends._win32com import Win32ComBackend

        _assets = Path(__file__).resolve().parent.parent / "assets"
        template = _assets / "test_template_readme_example.xlsx"

        data = [
            {"Date": datetime(2024, 1, 1), "Stock Price": 42.5},
            {"Date": datetime(2024, 2, 1), "Stock Price": 55.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table(
                "StockPriceHistory",
                data,
                FillXLConfig(copy_table_styles=True, extend_conditional_formatting=True),
            )
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["Stock History"]

                # Date column must remain a date format
                date_nf = ws.cell(row=12, column=2).number_format
                assert "yy" in date_nf.lower() or "dd" in date_nf.lower(), (
                    f"Date column format lost: {date_nf!r}"
                )

                # Stock Price column must be numeric, not a date
                price_val = ws.cell(row=12, column=3).value
                price_nf = ws.cell(row=12, column=3).number_format
                assert isinstance(price_val, (int, float)), (
                    f"Stock Price should be numeric, got {price_val!r}"
                )
                assert not ("yy" in price_nf.lower() or "dd" in price_nf.lower()), (
                    f"Stock Price column got date format: {price_nf!r}"
                )
                assert price_nf and "0" in price_nf, (
                    f"Stock Price format should be numeric: {price_nf!r}"
                )

                # Target column (formula) must still be present
                target_val = ws.cell(row=12, column=4).value
                assert target_val is not None, "Target formula missing"
                assert isinstance(target_val, str) and target_val.startswith("="), (
                    f"Target should be a formula: {target_val!r}"
                )
            finally:
                wb.close()
        finally:
            backend.close()

    def test_copy_table_styles_does_not_shift_formats(self, output_path: Path) -> None:
        """``copy_table_styles`` must copy from the data body columns only,
        not the full worksheet row, to avoid a rightward column shift.
        """
        from conftest import TEMPLATE

        from fillxl._backends._win32com import Win32ComBackend

        backend = Win32ComBackend()
        try:
            backend.load(TEMPLATE)
            data = [
                {"Col1": "a", "Col2": 1, "ColNotFilled": "x", "TestFormattedPct": 0.1},
                {"Col1": "b", "Col2": 2, "ColNotFilled": "y", "TestFormattedPct": 0.2},
                {"Col1": "c", "Col2": 3, "ColNotFilled": "z", "TestFormattedPct": 0.3},
            ]
            backend.fill_table("TableTest", data, FillXLConfig(copy_table_styles=True))
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["TestTable"]
                numeric_val = ws.cell(row=4, column=2).value
                numeric_nf = ws.cell(row=4, column=2).number_format
                assert isinstance(numeric_val, (int, float)), (
                    f"Col2 should be numeric: {numeric_val!r}"
                )
                # Must NOT have a date format (which would indicate a
                # column shift from the copy-paste)
                is_date_fmt = numeric_nf and (
                    "yy" in numeric_nf.lower() or "dd" in numeric_nf.lower()
                )
                assert not is_date_fmt, (
                    f"Col2 format should not be a date format, got {numeric_nf!r}"
                )
            finally:
                wb.close()
        finally:
            backend.close()

    def test_conditional_formatting_extension_preserves_formats(
        self,
        output_path: Path,
    ) -> None:
        """Extending conditional formatting must preserve per-column formats."""
        from fillxl._backends._win32com import Win32ComBackend

        _assets = Path(__file__).resolve().parent.parent / "assets"
        template = _assets / "test_template_readme_example.xlsx"

        data = [
            {"Date": datetime(2025, 1, 1), "Stock Price": 100.0},
            {"Date": datetime(2025, 2, 1), "Stock Price": 110.0},
        ]

        backend = Win32ComBackend()
        try:
            backend.load(template)
            backend.fill_table(
                "StockPriceHistory",
                data,
                FillXLConfig(
                    copy_table_styles=False,
                    extend_conditional_formatting=True,
                ),
            )
            backend.save(output_path)

            wb = openpyxl.load_workbook(output_path)
            try:
                ws = wb["Stock History"]
                val = ws.cell(row=12, column=3).value
                nf = ws.cell(row=12, column=3).number_format
                assert isinstance(val, (int, float)), f"Stock Price should be numeric: {val!r}"
                assert "0" in (nf or ""), f"Stock Price format should be numeric, got {nf!r}"
            finally:
                wb.close()
        finally:
            backend.close()
