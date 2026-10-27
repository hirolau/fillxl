# Engine parity

fillxl has two engines and one promise: **the same template and the same data
give the same workbook.** parity is checked by
`tests/parity/`, which fills a template with both engines and compares a
snapshot of every filled workbook (cell values, data types, number formats,
table refs, sort state, chart references, conditional-formatting ranges, sheet
states and defined names).

## How the engines are selected

| Template → output | `engine="auto"` | `openpyxl` | `win32com` |
|---|---|---|---|
| `.xlsx` → `.xlsx` | openpyxl | yes | yes |
| `.xlsm` → `.xlsm` | openpyxl (VBA kept) | yes | yes |
| `.xlsx` → `.xlsm` / `.xlsb` / `.xls` | Excel if available, else an error | refused | yes |
| `.xlsb` / `.xls` → anything | Excel if available, else an error | refused | yes |

A format the selected engine cannot write is refused **before the template is
loaded**, not after the fill has run.

## Feature support

`yes` means fillxl handles the feature; `differs` means the engines produce
different output and the difference is listed under *Known gaps*.

| Feature | openpyxl | win32com |
|---|---|---|
| Named cells (global and sheet-scoped) | yes | yes |
| Named cells on merged ranges | differs | differs |
| Tables that grow | yes | yes |
| Tables that shrink, totals rows included | yes | yes |
| Calculated (formula) columns | yes | yes |
| Table styles copied to new rows | yes | yes |
| Conditional formatting extended to the table | yes | yes |
| Charts: category, value, X/Y, bubble size, multi-level categories, custom error bars | yes | yes |
| Chart sheets | yes | yes |
| Pivot caches marked for refresh on open | yes | yes |
| Pivot tables refreshed immediately | no (open-to-refresh) | yes |
| Formula results recalculated before saving | no (recalc on open) | yes |
| Dates / datetimes formatting | yes | yes |
| Sheet hiding, deletion and activation | yes | yes |
| Lossy-template report (`on_lossy`) | yes | not needed |

## Known gaps

These are known differences or limitations. The "engine" column says where the
gap is; where both engines are named, they behave the same way.

| Gap | Engine | Status |
|---|---|---|
| `_xHHHH_` sequences in user text (`_x0041_`) are unescaped by Excel on the openpyxl path, because openpyxl writes them literally | openpyxl | open — escaping changes what openpyxl reads back, so it needs a decision |
| Data validation ranges are not extended when a table grows | both (openpyxl confirmed) | open — verify win32com first |
| Pivots and defined names that point at the table body with A1 references are not updated | both | open — exact matches will be updated in 0.8 |
| A defined name pointing at a merged range is refused as a multi-cell range | both | open |
| Slicers, timelines, threaded comments, form controls, ActiveX, the data model, Power Query parts, rich data, external links and Office add-ins are dropped by openpyxl | openpyxl | reported by `LossyTemplateWarning` |
| Modern charts (chartEx) are dropped by openpyxl | openpyxl | reported |
| Pictures are dropped by openpyxl when Pillow is not installed | openpyxl | reported |
| Chart style and colour parts (`style*.xml`, `colors*.xml`) are dropped by openpyxl | openpyxl | reported |
| `.xltx` / `.xltm` templates | openpyxl | refused; use `engine="win32com"` |
| A stale `sortState` ref is left as the template had it | both | not a defect: Excel 16 does the same on resize and opens the file without repairing it |
| Headerless tables | both | yes, but see the tests for the COM `Resize` quirks |

## Verifying a claim

Behaviour tests are parametrised over both engines (`tests/_engines.py`) and
skip the `win32com` case when Excel is not installed. Anything that genuinely
needs Excel — opening a file, repairing, recalculating — carries
`@pytest.mark.excel` and is skipped elsewhere:

```bash
pytest -m excel        # Windows with Microsoft Excel
pytest -m "not excel"  # anywhere
```

When a documented behaviour and Excel disagree, the test wins and the
documentation is wrong. Add a test that reproduces the difference before
changing either engine.
