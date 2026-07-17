# Excel Test File: test_template_2_charts.xlsx — Sheet Descriptions

10 chart sheets, each containing one Excel table bound to one chart.

---

## Sheet: 1-Column Chart

**Table:** `ColumnChartTable` — `A1:E5`

| Region | Q1     | Q2     | Q3     | Q4     |
|--------|--------|--------|--------|--------|
| North  | 45,000 | 52,000 | 48,000 | 61,000 |
| South  | 38,000 | 41,000 | 45,000 | 50,000 |
| East   | 52,000 | 58,000 | 55,000 | 63,000 |
| West   | 41,000 | 47,000 | 51,000 | 58,000 |

**Chart:** BarChart — one series per region, categories = Q1–Q4 (horizontal), values = quarterly amounts.

---

## Sheet: 2-Line Chart

**Table:** `LineChartTable` — `A1:C13`

Monthly data for 2025 with two value columns.

| Date       | 2024    | 2025    |
|------------|---------|---------|
| 2025-01-01 | 120,000 | 135,000 |
| 2025-02-01 | 115,000 | 142,000 |
| … 12 rows  | …       | …       |

**Chart:** LineChart — two series (2024, 2025), categories = Date column.

---

## Sheet: 3-Pie Chart

**Table:** `PieChartTable` — `A1:B7`

| Company  | Market Share |
|----------|-------------|
| Apple    | 28%         |
| Samsung  | 22%         |
| Xiaomi   | 14%         |
| Oppo     | 10%         |
| Vivo     | 8%          |
| Others   | 18%         |

**Chart:** PieChart — categories = Company, values = Market Share.

---

## Sheet: 4-Bar Chart

**Table:** `BarChartTable` — `A1:B9`

| Department        | Employees |
|-------------------|-----------|
| Engineering       | 145       |
| Sales             | 98        |
| Marketing         | 52        |
| Finance           | 38        |
| HR                | 25        |
| Operations        | 72        |
| Customer Support  | 64        |
| Legal             | 18        |

**Chart:** BarChart — categories = Department, values = Employees.

---

## Sheet: 5-Area Chart

**Table:** `AreaChartTable` — `A1:D9`

| Week    | Desktop | Mobile  | Tablet  |
|---------|---------|---------|---------|
| Week 1  | 45,000  | 62,000  | 12,000  |
| Week 2  | 48,000  | 68,000  | 13,500  |
| … 8 rows| …       | …       | …       |

**Chart:** AreaChart — three series (Desktop, Mobile, Tablet), categories = Week.

---

## Sheet: 6-Scatter Chart

**Table:** `ScatterChartTable` — `A1:C11`

| Product    | Price ($) | Units Sold |
|------------|-----------|------------|
| Product A  | 29        | 1,450      |
| Product B  | 49        | 980        |
| … 10 rows  | …         | …          |

**Chart:** ScatterChart — xVal = Price ($), yVal = Units Sold.

---

## Sheet: 7-Doughnut Chart

**Table:** `DoughnutChartTable` — `A1:B7`

| Category    | Budget   |
|-------------|----------|
| Salaries    | 450,000  |
| Marketing   | 120,000  |
| Operations  | 85,000   |
| Technology  | 95,000   |
| R&D         | 150,000  |
| Other       | 50,000   |

**Chart:** DoughnutChart — categories = Category, values = Budget.

---

## Sheet: 8-Stacked Column

**Table:** `StackedColumnChartTable` — `A1:E5`

| Quarter  | Electronics | Clothing | Food   | Home   |
|----------|-------------|----------|--------|--------|
| Q1 2025  | 85,000      | 42,000   | 38,000 | 25,000 |
| Q2 2025  | 92,000      | 48,000   | 41,000 | 28,000 |
| Q3 2025  | 88,000      | 55,000   | 45,000 | 32,000 |
| Q4 2025  | 105,000     | 68,000   | 52,000 | 38,000 |

**Chart:** BarChart (stacked) — four series (Electronics, Clothing, Food, Home), categories = Quarter.

---

## Sheet: 9-Combo Chart

**Table:** `ComboChartTable` — `A1:C13`

Monthly revenue and growth rate for 2026.

| Date       | Revenue  | Growth Rate |
|------------|----------|-------------|
| 2026-01-01 | 180,000  | 5%          |
| 2026-02-01 | 195,000  | 8%          |
| … 12 rows  | …        | …           |

**Chart:** BarChart — categories = Date, values = Revenue (first series, bars). Secondary axis series (Growth Rate, line) not exposed via openpyxl.

---

## Sheet: 10-Waterfall

**Table:** `WaterfallChartTable` — `A1:B9`

| Category           | Amount   |
|--------------------|----------|
| Starting Balance   | 500,000  |
| Product Sales      | 350,000  |
| Service Revenue    | 125,000  |
| Cost of Goods      | -180,000 |
| Operating Expenses | -145,000 |
| Marketing          | -75,000  |
| Taxes              | -62,000  |
| Ending Balance     | 513,000  |

**Chart:** Waterfall chart (openpyxl does not expose series references for this chart type). No chart series references require updating.
