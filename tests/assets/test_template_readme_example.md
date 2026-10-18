# Excel Test File: test_template_readme_example.xlsx — Sheet Descriptions

A stock-price tracker template used to demonstrate the basic `fillxl.fill()` workflow shown in the README.

---

## Sheet: Stock History

### Named Cells

| Name          | Address | Description                          |
|---------------|---------|--------------------------------------|
| `title`       | B2      | Report title (e.g. "ACME Corp")      |
| `currency`    | C4      | Currency label (e.g. "USD")          |
| `target_price`| C5      | Target price for the streak calc     |

### Other Cells

| Cell | Value                      | Description                                 |
|------|----------------------------|---------------------------------------------|
| B4   | `Currency:`                | Static label                                |
| B5   | `Target Price`             | Static label                                |
| B6   | `Max Consecutive days...`  | Static label                                |
| C6   | Array formula              | Calculates longest streak above target      |
| C7   | Array formula (START)      | First date of the longest streak            |
| C8   | Array formula (END)        | Last date of the longest streak             |
| B7   | `START`                    | Static label                                |
| B8   | `END`                      | Static label                                |

### Table: `StockPriceHistory` — `B10:E15`

| Column          | Description                            |
|-----------------|----------------------------------------|
| Date            | Monthly date (datetime)                |
| Stock Price     | Closing price (entered as string `=80` in row 11, numeric thereafter) |
| Target          | Formula referencing `target_price`     |
| Longest Streak  | Array formula for streak calculation   |

Placeholder data rows: 5 rows (2021-07 through 2021-11).

### Chart

**LineChart** — shows Stock Price over time. Bound to the `StockPriceHistory` table columns.

### Conditional Formatting

A single expression-based rule applied to `C11:E15`.

---

## Usage in the README

The README presents a simplified version of this pattern:

```python
fillxl.fill(
    "template.xlsx",
    "output.xlsx",
    tables={
        "SalesTable": [
            {"Product": "Widget", "Sales": 100},
        ],
    },
    values={
        "Title": "Q4 Report",
        "Date": "2024-01-01",
    },
)
```

The actual template (`test_template_readme_example.xlsx`) contains the richer stock-tracker scenario above, used for integration tests.
