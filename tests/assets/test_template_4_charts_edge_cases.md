# Excel Test File: test_template_4_charts_edge_cases.xlsx — Sheet Descriptions

Two sheets testing multiple charts on one sheet and a chart bound to a plain range (not a table).

---

## Sheet: 1-Multiple Column Charts

**Table:** `ColumnChartTable` — `A1:E5`

| Region | Q1     | Q2     | Q3     | Q4     |
|--------|--------|--------|--------|--------|
| North  | 45,000 | 52,000 | 48,000 | 61,000 |
| South  | 38,000 | 41,000 | 45,000 | 50,000 |
| East   | 52,000 | 58,000 | 55,000 | 63,000 |
| West   | 41,000 | 47,000 | 51,000 | 58,000 |

**Charts (3 on one sheet):**

1. **BarChart** — one series per region, categories = Q1–Q4 (horizontal), values = quarterly amounts. Positioned to the right of the table.
2. **LineChart** — same data as BarChart, placed further right.
3. **PieChart** — shows Region × Q1 values, positioned below the table.

This sheet tests that multiple chart types on a single worksheet all get their data references updated correctly after a table resize.

---

## Sheet: 2-Line Chart From Range

**No table.** This chart is bound to a plain A1:C13 range (not a ListObject).

| Date       | 2024    | 2025    |
|------------|---------|---------|
| 2025-01-01 | 120,000 | 135,000 |
| 2025-02-01 | 115,000 | 142,000 |
| 2025-03-01 | 130,000 | 148,000 |
| 2025-04-01 | 125,000 | 151,000 |
| 2025-05-01 | 140,000 | 158,000 |
| 2025-06-01 | 155,000 | 165,000 |
| 2025-07-01 | 148,000 | 172,000 |
| 2025-08-01 | 160,000 | 178,000 |
| 2025-09-01 | 165,000 | 185,000 |
| 2025-10-01 | 170,000 | 190,000 |
| 2025-11-01 | 180,000 | 198,000 |
| 2025-12-01 | 195,000 | 210,000 |

**Chart:** LineChart — two series (2024, 2025), categories = Date column.

This sheet tests that fillxl correctly handles charts bound to plain ranges rather than tables. Such charts should be preserved but **not** resized by fillxl (only table-bound chart references are updated).
