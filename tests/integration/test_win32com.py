"""Win32com integration tests (require real Excel via COM).

These tests verify that the win32com backend correctly interacts with
the Excel COM interface — lifecycle, named cells, sheet operations,
and table fills.  All tests are marked @pytest.mark.win32com.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from unittest import mock

import openpyxl
import pytest
from openpyxl.worksheet.table import (
    Table,
    TableColumn,
    TableFormula,
)

from fillxl import FillXLConfig


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


class TestReadCalculatedColumnFormulas:
    """Unit tests for ``_read_calculated_column_formulas``.

    These tests read workbooks with openpyxl (no Excel COM required).
    They verify that the helper correctly distinguishes real formulas
    from plain-number defaults, and that formula strings are returned
    with the ``=`` prefix.
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

    def test_real_formula_detected(self, tmp_path: Path) -> None:
        r"""A column with ``calculatedColumnFormula='Col1&\" \"&Col2'`` is
        detected as a formula column.
        """
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "real_formula.xlsx"
        self._make_table_wb(p, {"Col1": None, "ColFormula": 'Col1&" "&Col2'})

        cols, formulas = _read_calculated_column_formulas(p, "TestTable")
        assert "ColFormula" in cols
        assert formulas.get("ColFormula") == '=Col1&" "&Col2'

    def test_plain_number_skipped(self, tmp_path: Path) -> None:
        """A column with ``calculatedColumnFormula='100'`` (a plain number
        default) is NOT detected as a formula column.
        """
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "plain_number.xlsx"
        self._make_table_wb(p, {"Col1": None, "StockPrice": "100"})

        cols, _ = _read_calculated_column_formulas(p, "TestTable")
        assert "StockPrice" not in cols

    def test_plain_bool_skipped(self, tmp_path: Path) -> None:
        """A column with ``calculatedColumnFormula='TRUE'`` is skipped."""
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "plain_bool.xlsx"
        self._make_table_wb(p, {"Col1": None, "Flag": "TRUE"})

        cols, _ = _read_calculated_column_formulas(p, "TestTable")
        assert "Flag" not in cols

    def test_plain_float_skipped(self, tmp_path: Path) -> None:
        """A column with ``calculatedColumnFormula='12.5'`` is skipped."""
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "plain_float.xlsx"
        self._make_table_wb(p, {"Col1": None, "Price": "12.5"})

        cols, _ = _read_calculated_column_formulas(p, "TestTable")
        assert "Price" not in cols

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
            assert backend._read_calculated_column_formulas("TableA") == (
                ["Col1"],
                {"Col1": "=Col1+1"},
            )
            assert backend._read_calculated_column_formulas("TableB") == (
                ["Col1"],
                {"Col1": '=Col1&"!"'},
            )

        assert load_workbook.call_count == 1

    def test_formula_with_newlines_preserved_on_read(self, tmp_path: Path) -> None:
        r"""A formula with embedded newlines is returned with those newlines.

        Formatting is preserved end-to-end: the reader keeps indentation,
        and ``_prepare_formula_for_com`` only normalizes CR/CRLF and strips
        ``_xlfn.`` / ``_xlpm.`` prefixes.
        """
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "newline.xlsx"
        self._make_table_wb(p, {"Col1": None, "ColFormula": 'Col1&\n" "&Col2'})

        cols, formulas = _read_calculated_column_formulas(p, "TestTable")
        assert "ColFormula" in cols
        assert "\n" in formulas.get("ColFormula", "")

    def test_xlfn_prefix_preserved_in_raw_formula(self, tmp_path: Path) -> None:
        """A formula with ``_xlfn.`` prefix is returned as-is (the
        prefix stripping is done during formula *writing* in
        ``fill_table``).
        """
        from fillxl._backends._win32com import _read_calculated_column_formulas

        p = tmp_path / "xlfn.xlsx"
        self._make_table_wb(p, {"Col1": None, "ColFormula": "_xlfn.LET(x,1,x+1)"})

        cols, formulas = _read_calculated_column_formulas(p, "TestTable")
        assert "ColFormula" in cols
        assert formulas.get("ColFormula") == "=_xlfn.LET(x,1,x+1)"

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
