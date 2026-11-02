# Backends

fillxl supports three backend engines for writing Excel files. Choose based on your platform and fidelity requirements.

All engines write **the same values** for the same template and data: the same
cell values, number formats, chart references and sheet states. The differences
that remain — cached formula results, parts an engine cannot write — are listed
under *Feature support* below.

`engine="auto"` (the default) uses `openpyxl` for `.xlsx`/`.xlsm` templates
saved in the same format, and `win32com` when openpyxl cannot do the job
(`.xlsb`/`.xls` templates, format conversion) and Excel is available. A format
the selected engine cannot write is refused **before the template is loaded**,
not after the fill has run:

| Template → output | `auto` | `openpyxl` | `fillxl` | `win32com` |
|---|---|---|---|---|
| `.xlsx` → `.xlsx` | openpyxl | yes | yes | yes |
| `.xlsm` → `.xlsm` | openpyxl (VBA kept) | yes | yes (VBA kept) | yes |
| `.xlsx` → `.xlsm` / `.xlsb` / `.xls` | Excel if available, else an error | refused | refused | yes |
| `.xlsb` / `.xls` → anything | Excel if available, else an error | refused | refused | yes |

## Comparison

| Engine | Cross-platform | Requires Excel | Best for |
|--------|:---:|:---:|---|
| `openpyxl` | :material-check: | :material-close: | Most use cases, Linux/macOS |
| `fillxl` | :material-check: | :material-close: | Templates with spilled arrays, slicers, data types or other parts openpyxl cannot write |
| `win32com` | :material-close: Windows only | :material-check: | `.xlsm`/`.xlsb`, modern charts, full fidelity |

## Feature support

`yes` means fillxl handles the feature; `differs` means the engines behave
differently and the difference is described below.

| Feature | openpyxl | fillxl | win32com |
|---|---|---|---|
| Named cells, global and sheet-scoped | yes | yes | yes |
| Tables that grow and shrink, totals rows included | yes | yes | yes |
| Calculated (formula) columns | yes | yes | yes |
| Table styles copied to new rows | yes | yes | yes |
| Conditional formatting extended to the table | yes | yes | yes |
| Charts: categories, values, X/Y, bubble sizes, multi-level categories, custom error bars | yes | yes | yes |
| Chart sheets | yes | yes | yes |
| Pivot caches marked for refresh on open | yes | yes | yes |
| Pivot tables refreshed immediately | differs — open-to-refresh | differs — open-to-refresh | yes |
| Formula results recalculated before saving | differs — recalc on open | differs — recalc on open | yes |
| Cached results on the sheet a fill writes | dropped | dropped | computed |
| Summary formulas and A1 totals formulas follow a resize | yes | yes | yes (Excel does it) |
| Pivot sources, defined names, validation lists follow a resize | warns — `UntrackedRangeWarning` | warns — `UntrackedRangeWarning` | yes (Excel does it) |
| Hidden rows shown and table filters cleared by a fill | yes | yes | yes |
| Sheet hiding, deletion and activation | yes | yes | yes |
| Spilled dynamic arrays (`SEQUENCE`, `FILTER`, `SORT`, …) and cell metadata | differs — lost | yes | yes |
| Slicers, timelines, `chartEx`, rich data, VBA, embedded objects | differs — lost | yes | yes |
| Lossy-template report (`on_lossy`) | yes | not needed | not needed |

How the differences are checked: the parity suite fills one template with every
engine and compares a snapshot of the result — cell values, data types,
number formats, table refs, sort state, chart references, conditional
formatting ranges, sheet states and defined names. See
[Development](development.md) for how to run it, and [Benchmarks](benchmarks.md)
for how the three engines compare on the same fill.

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
- No dynamic array formula support: a template with spilled ranges loses both
  the spill markup and `xl/metadata.xml`. Use `engine="fillxl"` to keep them
- Writes no cached formula results: the output sets `fullCalcOnLoad` so Excel
  recalculates on open. Use `engine="win32com"` when the saved file must
  already contain computed values.
- Pivot sources, defined names and validation lists that cover a table body are
  not resized; the fill warns with `UntrackedRangeWarning` instead.
- Deleting a sheet leaves formulas, charts, names and pivots that refer to it
  pointing at a sheet that no longer exists, and warns
  (`DanglingReferenceWarning`) rather than rewriting them.

---

## fillxl

Edits the OOXML package itself: pure Python, no Excel installation, and no
object model to lose features in.

```python
fill(template="tpl.xlsx", output="out.xlsx", ..., engine="fillxl")
```

The template is a ZIP of XML parts. Only the parts a fill has to change are
parsed and rewritten — the worksheet, the table definition, chart references,
pivot cache flags and `xl/workbook.xml`. Every other part is copied byte for
byte, so a feature the engine does not model cannot disappear: spilled
dynamic arrays and their `xl/metadata.xml`, `cm`/`vm` cell metadata, slicers,
`chartEx`, rich data, embedded objects and VBA all survive.

### Pros

- Works on **Linux, macOS, and Windows**; nothing to install beyond fillxl
- Keeps every template feature, including spilled arrays — `on_lossy` stays
  silent because nothing is lost
- Fills `.xlsx` and macro-enabled `.xlsm` templates (VBA is preserved)

### Limitations

- Reads and writes `.xlsx`/`.xlsm` only, and only in the same format: `.xlsb`,
  `.xls`, encrypted workbooks and format conversion stay on `win32com`
- Modern charts (`chartEx`) are left untouched, so they do not follow a table
  resize
- Writes no cached formula results, and drops the ones the template stored on
  the sheet a fill writes; the output asks Excel to recalculate on open, like
  the openpyxl engine
- Pivot sources, defined names and validation lists that cover a table body are
  not resized — it warns with `UntrackedRangeWarning` instead
- Deleting a sheet leaves formulas, charts, names and pivots that refer to it
  pointing at a sheet that no longer exists, and warns
  (`DanglingReferenceWarning`) rather than rewriting them

### How it stays safe

Rows are never inserted or removed — a growing table writes into cells verified
empty — and an array (spill) range counts as occupied even when its cells look
empty. A save parses every part the fill rewrote before the output file is
opened, so a bad edit raises `RuntimeError` instead of producing a workbook no
reader can open. `on_lossy` acts as an assertion here: if anything at all were
missing from the output, the fill would say so. [The fillxl
engine](fillxl-engine.md) explains how it works and where its limits are.

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

Every engine writes non-empty strings as text, even when Excel would otherwise
re-parse them (`"00123"` stays `"00123"`, `"1/2"` stays a string). This is
handled by prefixing strings with Excel's text marker when they are written
through COM; the openpyxl engine uses the text cell type, including for
error-like strings such as `"#N/A"`. Empty strings become empty cells.

Strings starting with `=` stay literal in every engine by default, so
untrusted data cannot inject formulas. Pass
`FillXLConfig(escape_formulas=False)` when the input is trusted and a leading
`=` is meant to create a formula.

Excel tables always end up with **at least one body row**: filling zero rows
clears and keeps a single row, because Excel rejects a table without a body.

Pivot caches are refreshed by every engine: openpyxl and fillxl mark them
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

### Formulas the fill rewrites

Excel edits formulas when a table changes height; every engine reproduces
that, so the same template and data give the same numbers:

- **Calculated columns.** A column whose `calculatedColumnFormula` uses A1
  references (`=B5*C5`) is written row by row — `=B5*C5`, `=B6*C6`, `=B7*C7` —
  because a relative row moves with the copy while a `$`-anchored row
  (`=B5*$C$5`, `=B5*C$5`) does not. Structured references
  (`=[@Qty]*[@Price]`) already resolve per row and are written unchanged.
- **Formulas next to the table.** A worksheet formula whose range covers the
  table body row for row — `=SUM(B4:B4)` for a one-row body — is resized with
  the table (`=SUM(B4:B7)`), including `$`-anchored ranges, ranges on other
  sheets that name the table's sheet, and the same range inside a larger
  formula. Single-cell references (`=B4*2`), ranges that start or end outside
  the body (`=SUM(B4:B5)`, `=SUM(Z1:Z1)`), string literals that merely look
  like ranges, and references to other sheets are left exactly as they are.

Fillxl only rewrites what it can attribute to the table. A formula that reads
table rows without covering the body exactly (for example
`=SUM(B4:B5)` when the body is one row) is left alone, so it will not follow a
resize — Excel would leave it alone too.

### Losing template features (openpyxl)

The openpyxl engine cannot preserve every Excel feature. After saving to the
staging file (and before the atomic swap) it compares feature counts in the
template and the output: modern charts, slicers, timelines, shapes, pictures,
embedded objects, rich data, VBA, external links and worksheet extensions.
`FillXLConfig.on_lossy` controls the result:

| Value | Behaviour |
|-------|-----------|
| `"ignore"` | Do nothing. |
| `"warn"` (default) | Issue a `LossyTemplateWarning` listing the lost features. |
| `"raise"` | Raise `LossyTemplateError`; because the swap has not happened, the previous output file is left untouched. |

[Warnings](warnings.md) lists this and every other category fillxl raises.

The detector compares **part and element counts**. It catches dropped parts,
shapes, worksheet extensions and `xl/metadata.xml` (spilled arrays and data
types), but it cannot prove that a preserved part is byte-for-byte unchanged. Parts that only a sheet removed by `delete_sheets`
referenced are ignored: their disappearance is the caller's doing, not the
engine's.

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
