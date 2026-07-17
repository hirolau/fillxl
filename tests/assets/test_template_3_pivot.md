# Excel Test File: test_template_3_pivot.xlsx — Sheet Descriptions

---

## Sheet: Sales Data

**Table:** `SalesData` — `A1:E19`

Source data for the pivot table. 18 rows of sales records.

| Region | Product | Sales Rep | Quarter | Amount |
|--------|---------|-----------|---------|--------|
| North  | Widgets | Alice     | Q1      | 12,500 |
| North  | Gadgets | Alice     | Q1      | 8,700  |
| North  | Widgets | Bob       | Q1      | 9,800  |
| South  | Widgets | Carol     | Q1      | 11,200 |
| South  | Gadgets | Carol     | Q1      | 7,500  |
| East   | Widgets | Dave      | Q1      | 10,300 |
| East   | Gadgets | Dave      | Q1      | 6,800  |
| West   | Widgets | Eve       | Q1      | 13,100 |
| West   | Gadgets | Eve       | Q1      | 9,200  |
| North  | Widgets | Alice     | Q2      | 14,200 |
| North  | Gadgets | Alice     | Q2      | 9,100  |
| North  | Widgets | Bob       | Q2      | 10,500 |
| South  | Widgets | Carol     | Q2      | 12,800 |
| South  | Gadgets | Carol     | Q2      | 8,200  |
| East   | Widgets | Dave      | Q2      | 11,700 |
| East   | Gadgets | Dave      | Q2      | 7,400  |
| West   | Widgets | Eve       | Q2      | 15,000 |
| West   | Gadgets | Eve       | Q2      | 10,100 |

---

## Sheet: Pivot Analysis

**Pivot Table:** `SalesPivot`

Based on `SalesData` table (cache source = `'Sales Data'!$A$1:$E$19`).

refreshOnLoad = `None` (not set — default behaviour)

Layout:

| Row Labels | Gadgets   | Widgets   | Grand Total |
|------------|-----------|-----------|-------------|
| East       | 14,200    | 22,000    | 36,200      |
| North      | 17,800    | 47,000    | 64,800      |
| South      | 15,700    | 24,000    | 39,700      |
| West       | 19,300    | 28,100    | 47,400      |
| Grand Total| 67,000    | 121,100   | 188,100     |

Row fields: Region
Column fields: Product
Values: Sum of Amount
