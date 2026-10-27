# Backends

fillxl supports two backend engines for writing Excel files. Choose based on your platform and fidelity requirements.

Both engines produce **the same workbook** for the same template and data: the
same values, number formats, chart references and sheet states. The differences
that remain are listed under *Feature support* below.

`engine="auto"` (the default) uses `openpyxl` for `.xlsx`/`.xlsm` templates
saved in the same format, and `win32com` when openpyxl cannot do the job
(`.xlsb`/`.xls` templates, format conversion) and Excel is available. A format
the selected engine cannot write is refused **before the template is loaded**,
not after the fill has run:

| Template → output | `auto` | `openpyxl` | `win32com` |
|---|---|---|---|
| `.xlsx` → `.xlsx` | openpyxl | yes | yes |
| `.xlsm` → `.xlsm` | openpyxl (VBA kept) | yes | yes |
| `.xlsx` → `.xlsm` / `.xlsb` / `.xls` | Excel if available, else an error | refused | yes |
| `.xlsb` / `.xls` → anything | Excel if available, else an error | refused | yes |

## Comparison

| Engine | Cross-platform | Requires Excel | Best for |
|--------|:---:|:---:|---|
| `openpyxl` | :material-check: | :material-close: | Most use cases, Linux/macOS |
| `win32com` | :material-close: Windows only | :material-check: | `.xlsm`/`.xlsb`, modern charts, full fidelity |

## Feature support

`yes` means fillxl handles the feature; `differs` means the engines behave
differently and the difference is described below.

| Feature | openpyxl | win32com |
|---|---|---|
| Named cells, global and sheet-scoped | yes | yes |
| Tables that grow and shrink, totals rows included | yes | yes |
| Calculated (formula) columns | yes | yes |
| Table styles copied to new rows | yes | yes |
| Conditional formatting extended to the table | yes | yes |
| Charts: categories, values, X/Y, bubble sizes, multi-level categories, custom error bars | yes | yes |
| Chart sheets | yes | yes |
| Pivot caches marked for refresh on open | yes | yes |
| Pivot tables refreshed immediately | differs — open-to-refresh | yes |
| Formula results recalculated before saving | differs — recalc on open | yes |
| Sheet hiding, deletion and activation | yes | yes |
| Lossy-template report (`on_lossy`) | yes | not needed |

How the differences are checked: the parity suite fills one template with both
engines and compares a snapshot of the result — cell values, data types,
number formats, table refs, sort state, chart references, conditional
formatting ranges, sheet states and defined names. See
[Development](development.md) for how to run it.

---

## openpyxl (default)

The default backend. Pure Python, no Excel installation required.

```python
fill(template="tpl.xlsx", output="out.xlsx", ..., engine="openpyxl")
```

### Pros

- Works on **Linux, macOS, and Windows**
- No Excel license needed
- Fast and lightweight
- Fills `.xlsx` and macro-enabled `.xlsm` templates (VBA is preserved)

### Limitations

- Does **not** support `.xlsb` or legacy `.xls`
- Drops some modern chart types (e.g., sunburst, treemap), slicers and other
  parts; `on_lossy` reports what was lost as `LossyTemplateWarning`
- No dynamic array formula support
- Writes no cached formula results: the output sets `fullCalcOnLoad` so Excel
  recalculates on open. Use `engine="win32com"` when the saved file must
  already contain computed values.

---

## win32com

Uses the Excel COM API for maximum fidelity. Requires Windows with Excel installed.

```python
fill(template="tpl.xlsm", output="out.xlsm", ..., engine="win32com")
```

### Pros

- **Full fidelity** — all Excel features preserved
- Supports all modern chart types
- Handles `.xlsm` and `.xlsb` formats
- Recalculates before refreshing pivot caches and before saving, so pivots on
  calculated columns and formulas reading named cells show the new data

### Limitations

- Windows only
- Requires a licensed Excel installation
- Slower per-call overhead (Excel COM startup)

### Security

Before any workbook opens, fillxl sets Excel's `AutomationSecurity` to
force-disable macros, turns events off (`EnableEvents = False`) and stops link
updates (`AskToUpdateLinks = False`; workbooks are opened with
`UpdateLinks=0`). Failing to apply these settings is fatal. The backend also
never uses the clipboard: styles are copied with `AutoFill(Type=xlFillFormats)`
and conditional-formatting ranges are resized with `ModifyAppliesToRange`, so
whatever the user had on the clipboard survives a fill.

### Output format

The save format is inferred from the output suffix, so macros survive:

| Suffix | Excel format |
|--------|--------------|
| `.xlsx` | macro-free workbook |
| `.xlsm` | macro-enabled workbook (keeps `vbaProject.bin`) |
| `.xlsb` | binary workbook |
| `.xls` | legacy workbook |

Any other suffix raises `ValueError` instead of silently writing a different
format.

---

## Engine-specific behaviour worth knowing

Both engines write non-empty strings as text, even when Excel would otherwise
re-parse them (`"00123"` stays `"00123"`, `"1/2"` stays a string). This is
handled by prefixing strings with Excel's text marker when they are written
through COM; the openpyxl engine uses the text cell type, including for
error-like strings such as `"#N/A"`. Empty strings become empty cells in both
engines.

Strings starting with `=` stay literal in both engines by default, so
untrusted data cannot inject formulas. Pass
`FillXLConfig(escape_formulas=False)` when the input is trusted and a leading
`=` is meant to create a formula.

Excel tables always end up with **at least one body row**: filling zero rows
clears and keeps a single row, because Excel rejects a table without a body.

Pivot caches are refreshed by both engines: openpyxl marks them
(`refreshOnLoad`) so Excel recalculates on open, and win32com recalculates them
during the fill as well.

Tables with a **totals row** are supported: the data body is replaced and the
totals row is moved to remain directly below it, keeping its values, labels,
functions and styling. The table reference includes the totals row while the
auto-filter range stops at the last data row, matching what Excel writes.

`copy_table_styles=False` leaves new rows unformatted on the openpyxl engine.
Excel itself extends a table column's number format to new rows when the table
grows through COM, so a `win32com` fill can still show the column format even
with style copying disabled. This divergence is intentional.

openpyxl does **not** enforce worksheet protection when it writes a cell;
Excel does. A named cell on a protected sheet therefore fills with the
openpyxl engine but raises `RuntimeError` through win32com. This divergence is
intentional.

### Losing template features (openpyxl)

The openpyxl engine cannot preserve every Excel feature. After saving to the
staging file (and before the atomic swap) it compares feature counts in the
template and the output: modern charts, slicers, timelines, shapes, rich data,
VBA, external links and worksheet extensions. `FillXLConfig.on_lossy` controls
the result:

| Value | Behaviour |
|-------|-----------|
| `"ignore"` | Do nothing. |
| `"warn"` (default) | Issue a `LossyTemplateWarning` listing the lost features. |
| `"raise"` | Raise `LossyTemplateError`; because the swap has not happened, the previous output file is left untouched. |

[Warnings](warnings.md) lists this and every other category fillxl raises.

The detector compares **part and element counts**. It catches dropped parts,
shapes and worksheet extensions, but it cannot prove that a preserved part is
byte-for-byte unchanged.

Tables **without a header row** (the table's "Header Row" option is off) are
supported too. Excel names their columns `Column1`, `Column2`, …, so that is
how the data is addressed:

```python
fill(
    template="test.xlsx",
    output="TestFilled.xlsx",
    tables={"Table1": [{"Column1": 1, "Column2": 6}]},
)
```

The data starts in the table's first row (there is no header to skip). The
win32com engine writes these tables row by row, because Excel removes the
table when a whole headerless body is assigned in a single call.

!!! note "Totals-row formulas"
    Totals formulas that use structured references — for example
    `SUBTOTAL(109,Table1[Value])`, which is what Excel writes for its built-in
    totals functions — keep working because they follow the table. A formula
    typed with an explicit range (for example `SUM(B2:B5)`) is preserved
    verbatim and is **not** rewritten when the table grows or shrinks; use the
    built-in totals function or a structured reference instead.

!!! warning "CI/CD"
    win32com tests are excluded from CI (Linux runners can't run Excel COM).
    This backend is tested manually on Windows before each release.
