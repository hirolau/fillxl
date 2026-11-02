# Backend research: filling Excel without Excel and without openpyxl

**Status:** research only — no production code was changed.
**Date:** 2026-11-01
**Scope:** (1) a new engine that edits the OOXML package directly; (2) a new engine built on
[excelize-py](https://github.com/xuri/excelize-py).

---

## 0. TL;DR

| | Direct OOXML editing ("xmledit") | excelize-py |
|---|---|---|
| Needs Excel | no | no |
| Needs openpyxl | no | no |
| Keeps dynamic arrays / cell metadata | **yes — verified end to end with spilled arrays (§3.5)** | yes (measured) |
| Table resized with totals row | yes | **no** — `totalsRowCount` is zeroed by excelize (`adjust.go:692`) |
| Calculated columns kept | yes | **no** — `xlsxTableColumn` has no `calculatedColumnFormula` field |
| Chart ranges updated after resize | yes | **no** — no chart API (`add_chart`/`delete_chart` only) |
| Conditional formatting extended | yes | **no** — no getter for existing rules |
| Pivot caches marked for refresh | yes | **no** — no setter for existing caches |
| `.xls` / `.xlsb` | no | no |
| Encrypted `.xlsx` (password) | no | **yes** (measured) |
| New runtime dependency | none (stdlib only) | 29 MB compiled wheel, one maintainer |
| Effort | ~4–6 weeks to parity + hardening | ~1 week wiring, but 6 blockers that can only be fixed by post-patching the zip — i.e. by writing the XML editor anyway |

**Recommendation:** build the direct-OOXML engine. It is the only option that satisfies
"no Excel, no openpyxl" *and* keeps the features fillxl already promises
(totals rows, calculated columns, chart ranges, pivot refresh, conditional formatting).
Spilled dynamic arrays were verified end to end against Excel: a moved spill keeps its values and a
spill whose size depends on the table is recalculated and grows on open — provided the editor also
**deletes `xl/calcChain.xml`** and **shifts `<f t="array" ref>`**, both of which are hard requirements
(§3.5). Keep `win32com` for `.xls`/`.xlsb` and format conversion. excelize is not a viable primary
engine for fillxl's table-resize workload; at best it is a future fallback for reading
encrypted or very large workbooks.

Everything in the tables above was measured on this machine (Excel 16.0 via COM) or read in
the library source; claims that were not verified are marked as such. The raw evidence is in
[§8](#8-appendix--evidence-and-reproduction).

---

## 1. What a fillxl backend must actually do

The backend seam is [`Backend`](src/fillxl/_backend.py) — nine methods:

| Method | Used by | Notes |
|---|---|---|
| `load(path)` | `_core.fill` | copies the template to a temp dir first |
| `fill_named_cell(name, value, config)` | `_cells.populate_named_cells` | global + sheet-scoped names |
| `fill_table(name, data, config) → TableChangeSpecs` | `_tables.populate_table`, `_plan.plan_table_fill` | grow/shrink, totals rows, calculated columns, style copy, conditional-format extension |
| `update_charts(spec)` | `_charts.update_chart_data_ranges` | rewrites the series reference formulas and the point-count cache |
| `set_pivot_refresh()` | pivot caches | `refreshOnLoad="1"` |
| `hide_sheets(prefix)`, `delete_sheets(names)`, `set_active_sheet(name)` | `_sheets`, `_dangling` | must never leave zero visible sheets |
| `save(path)`, `close()` | `_core` | staging file + atomic replace |
| `check_output(template, output, policy)` | `_lossy.check_lossy` | report/raise on dropped features |

Non-obvious requirements that come out of the shared, backend-neutral code:

* **`_plan.py`** already computes every row span (`overflow_rows`, `unused_body_rows`,
  `vacated_rows`, `new_ref`, `new_autofilter_ref`) without touching a workbook, so a new
  backend inherits the tricky arithmetic including totals rows and "empty data still leaves one
  body row".
* **`_charts.py`** already contains the reference-rewriting logic (`update_reference_text`) as
  *text* operations on sheet references; only iteration over the chart objects is openpyxl-specific.
* **`_formulas.resolve_formula`** decides which table columns are calculated columns; the evidence
  gathering (`formula_definitions`, `classify_body`) is separated from the openpyxl accessors.
* **Overflow safety** ([`_tables._check_overflow`](src/fillxl/_tables.py)) refuses to grow a table
  into occupied cells. This is a *contract*, and it is what makes a surgical XML editor tractable:
  the backend never has to shift arbitrary user content out of the way.

Coupling inventory (openpyxl imports in `src/`):

| Module | openpyxl coupling | Reusable for a new backend |
|---|---|---|
| `_backend.py`, `_core.py`, `_plan.py`, `_charts.py`, `_formulas.py`, `_formats.py` | `CellRange`, `get_column_letter`, `is_date_format` only | yes — near-pure logic |
| `_normalize.py` | `ILLEGAL_CHARACTERS_RE` only | yes |
| `_adapters.py`, `_validation.py`, `_sheets.py`, `_lossy.py`, `_warnings.py` | none | yes, unchanged |
| `_cells.py`, `_tables.py`, `_dangling.py`, `_utils.py` | deep (worksheet objects, `_cells`, `Table`) | need a backend-neutral model |
| `_backends/_openpyxl.py` | all of it | replaced |

---

## 2. Method and environment

Measurements taken on this machine:

| | |
|---|---|
| OS | Windows (same box as fillxl development) |
| Excel | 16.0, reachable over COM |
| Python | 3.13.3 in a throw-away venv (plus the repo's 3.12 venv) |
| excelize | `excelize==0.1.0` (PyPI wheel, 29.2 MB, `py3-none-any`), bundling Go `excelize v2.11.1-0.20261011032807` |
| openpyxl | 3.1.5 |
| fillxl | current `main` working tree |

Technique: build synthetic templates, drive each candidate backend over them, then (a) diff the
OOXML package part by part and (b) open every result in real Excel and read back tables, totals
rows, formulas and chart series. Excel refusing to open a file is treated as a hard failure —
it is exactly what a user would see. A second round (§3.5) used templates authored **by Excel
itself** so that spilled dynamic arrays carry authentic markup, and edited those packages with the
direct-XML approach alone.

---

## 3. Option A — a backend that edits the OOXML package itself

### 3.1 The idea

`.xlsx`/`.xlsm` is a ZIP of XML parts. openpyxl's problem is not that it reads XML badly; it is
that it **rebuilds the whole package from its object model**, so anything the model cannot
represent (slicers, `chartEx`, `richData`, `metadata.xml`, …) disappears. fillxl's
[`_lossy.py`](src/fillxl/_lossy.py) exists purely to report that damage.

A surgical editor inverts the model: **copy every ZIP entry byte-for-byte, parse only the parts
that must change, and write those back.** Untouched features cannot be lost, because they are
never touched.

### 3.2 What each fillxl feature means in XML

| Feature | Parts | Edit |
|---|---|---|
| named cell write | `xl/workbook.xml` `<definedNames>` → target part | resolve `localSheetId`/name to `Sheet!$A$1`, then write one cell |
| cell write | `xl/worksheets/sheetN.xml` | `<c r="C3" s="2"><v>…</v></c>`; text as `t="inlineStr"` or appended `sst` entry; `t="b"`; `t="e"`; `<f>` for formulas |
| style copy into new rows | — | copy the `s="…"` attribute from the template's first data row; `xl/styles.xml` only needs touching when a **date** lands in a non-date cell (`_formats.py` already defines the rule) |
| table resize | `xl/tables/tableN.xml` | `ref`, `<autoFilter ref>`, per-column `calculatedColumnFormula`/`totalsRow*` (kept as-is) |
| grow/shrink rows | `xl/worksheets/sheetN.xml` | insert/remove `<row>` elements; renumber `r=` on rows/cells below; shift the totals row |
| chart ranges | `xl/charts/chartN.xml` | rewrite `<f>` texts (reuse `_charts.update_reference_text`) and `<ptCount>` |
| pivot refresh | `xl/pivotCache/pivotCacheDefinitionN.xml` | add `refreshOnLoad="1"` |
| conditional-format extension | sheet `<conditionalFormatting sqref>` | rewrite `sqref` (reuse the tested range logic in `_tables._cf_range_tracks_table`) |
| hide/delete/activate sheet | `xl/workbook.xml` `<sheets>`, `bookViews` | `state="hidden"`, drop the `<sheet>`, its part, its rels and content-type override, fix `activeTab`/`tabSelected` |
| recalc on open | `xl/workbook.xml` `<calcPr>` | set `fullCalcOnLoad="1"`; **delete `xl/calcChain.xml`** (part + `<Override>` + workbook relationship) — Excel rebuilds it |
| spilled arrays / dynamic formulas | `xl/worksheets/sheetN.xml` (`<f t="array" ref>`, `cm`), `xl/metadata.xml` | keep `metadata.xml` and `cm` untouched; move the `ref` with the master cell; refuse to insert *inside* a range (§3.5) |

### 3.3 The parts that are genuinely hard

1. **Row shifting.** Growing a table means inserting `<row>` elements and moving everything below.
   The prototype shows this is mechanical *if* you also shift: `<mergeCell ref>`, `<conditionalFormatting sqref>`,
   `<dataValidation sqref>`, `<hyperlink ref>`, drawing anchors (`xdr:from`/`xdr:to` rows),
   other tables' `ref`s, defined-name texts, chart references **and array/spill `ref`s**. fillxl's
   existing overflow refusal means only the totals row (and deliberately placed content below) is
   ever in the way, which bounds the problem — but the general case must still be implemented, or
   documented.
2. **`xl/calcChain.xml` blocks the file once rows move (measured).** Any row insertion makes the
   chain stale, and a stale chain is not repaired — Excel **refuses to open the workbook**. Removing
   the whole part (part + `<Override>` + workbook relationship) is enough; Excel rebuilds it. Same
   edit, with the chain kept: `Open FAILED`; with the chain removed: `OPEN OK` (§3.5, §8).
3. **Namespace prefixes.** `xml.etree.ElementTree` **renames prefixes** (`ns0:worksheet`) and drops
   comments unless `ET.register_namespace` is called for every prefix — verified. That is unsafe
   for `mc:Ignorable="x14ac xr xr2 xr3"`, which refers to prefixes by name. Two workable choices:
   * register the known prefix table (main, `r`, `mc`, `x14ac`, `x14`, `xr`, `xr2`, `xr3`, `a`, `c`,
     `cx`, `xdr`, …) and **refuse to rewrite** a part that uses an unregistered prefix, or
   * add `lxml` (manylinux/macOS/Windows wheels exist) which round-trips prefixes, comments and
     CDATA faithfully. That is a compiled dependency — the point of this engine is to have none,
     so preference is stdlib + prefix registry + a conservative fallback to byte-level patching.
4. **Byte-level patching vs. parsing.** Attribute edits (`ref="B2:D8"`, `refreshOnLoad`) are safest as
   literal text replacements; structural edits (row insertion) need parsing. A hybrid keeps the
   blast radius small, at the cost of writing two carefully tested code paths.
5. **Shared strings vs inline strings.** Either append to `xl/sharedStrings.xml` (must maintain
   `count`/`uniqueCount`, and must not renumber existing entries) or write `t="inlineStr"`, which
   never touches another part. Verified in Excel: inline strings added next to an existing
   shared-strings table are read correctly (`B7 = "inline-added"` alongside `B3 = "North"` from the
   `sst`). Recommendation: reuse an existing shared string when present, otherwise inline — simpler
   and provably non-destructive.
6. **Binary formats.** `.xlsb` (binary OOXML) and `.xls` (BIFF8) are out of scope; `win32com`
   stays for those. Encrypted `.xlsx` is a CFB/ECMA-376-agile container, not a zip — also out of
   scope (excelize handles it, see §4).

### 3.4 Evidence that it works

A throw-away prototype (~120 lines, no openpyxl, no Excel) grew a template table from 2 to 5 data
rows. On the first template used (15 parts: table, chart, conditional formatting, data validation,
hyperlink, merge, rich text) the part accounting was:

```
parts preserved byte-for-byte : docProps/*, xl/styles.xml, xl/theme/theme1.xml, xl/drawings/*,
                                xl/worksheets/sheet1.xml.rels, [Content_Types].xml, _rels/.rels  (11 of 15)
parts rewritten               : xl/worksheets/sheet1.xml, xl/tables/table1.xml,
                                xl/charts/chart1.xml, xl/workbook.xml  (4 of 15)
```

The Excel-validated run used a template with a totals row, a calculated column, a chart and
conditional formatting. Package state after the same grow:

```
table after edit : ref="B2:D8" totalsRowCount="1"
                   <calculatedColumnFormula>C3*2</calculatedColumnFormula>   (kept)
                   <autoFilter ref="B2:D7"/>                               (totals row excluded)
chart after edit : SERIES(Data!$C$2,Data!$B$3:$B$7,Data!$C$3:$C$7,1)
```

Opened in Excel 16:

```
final_prototype.xlsx  OK  table=$B$2:$D$8 totals=$B$8:$D$8
    Double D3==C3*2
    series==SERIES(Data!$C$2,Data!$B$3:$B$7,Data!$C$3:$C$7,1)
    rows3-9=[('North',10.0,20.0),('South',20.0,40.0),('East',30.0,60.0),
             ('West',40.0,80.0),('North',55.0,110.0),('Total',155.0,None),(None,None,None)]
```

Same template, same grow operation through excelize (see §4.3) produced a file **Excel refused to open**.

`find_lost_features(template, prototype_output)` (fillxl's own `on_lossy` detector) returns
`[]` — i.e. under this engine `on_lossy` becomes an assertion that nothing was dropped, instead of
a warning generator.

### 3.5 Spilled arrays and dynamic formulas — verified end to end

Spilled (`SEQUENCE`, `SORT`, `UNIQUE`, `FILTER`, …) formulas are the feature openpyxl is most often
blamed for losing, so the OOXML editor was verified against Excel-native dynamic arrays
specifically. The template was authored **by Excel itself** (COM), so the markup is authentic:

```
<c r="F8" cm="1"><f t="array" ref="F8:F12">_xlfn.SEQUENCE(5)</f><v>1</v></c>
<c r="F9"><v>2</v></c> … <c r="F12"><v>5</v></c>          spill children carry cached values
<c r="H8"><f>SUM(_xlfn.ANCHORARRAY(F8))</f><v>15</v></c>   the "F8#" operator
<c r="J3" t="str" cm="1"><f t="array" ref="J3:J4">…SORT(UNIQUE(Tbl[Region]))</f><v>North</v></c>
xl/metadata.xml: metadataType XLDAPR + cellMetadata count="1"   ← what "dynamic" actually means
xl/calcChain.xml: <c r="F8" i="1" a="1"/> …
```

Three rules came out of the experiment. The first two are hard requirements; the third is a safety
rule for the fill path.

**Rule 1 — delete `xl/calcChain.xml` whenever rows move.** A stale chain is not repaired by Excel, it
blocks the file. Same edit, chain kept vs. removed:

| file | edit | calcChain | Excel |
|---|---|---|---|
| `A_refs.xlsx` | rows + array `ref` shifted correctly | kept (stale) | **OPEN FAILED** |
| `A_refs_nochain.xlsx` | identical edit | removed | **OPEN OK** |
| `A_naive.xlsx` | rows shifted, array `ref` left stale | kept | **OPEN FAILED** |
| `A_naive_nochain.xlsx` | identical edit | removed | **OPEN FAILED** ← stale `ref` is fatal on its own |

**Rule 2 — move `<f t="array" ref>` with its master cell, and refuse to insert inside a range.**
`A_naive_nochain` proves a `ref` that no longer matches the cell is fatal, so the `ref` must be
shifted. Conversely an insertion point that falls *strictly inside* a spill range cannot be expressed
at all (Excel itself blocks it as "you can't change part of an array"), so the editor must raise
instead of producing a half-shifted range — this is the "straddling" check in the prototype
(`array/spill range F4:F11 straddles row 5; refusing to insert`).

**Rule 3 — treat an array's `ref` range as occupied**, not just cells that hold values. Spill
children normally carry `<v>`, so a value-based overflow check already refuses to grow a table into a
spill — but a spill whose results are blank writes `<c r="C8" t="str"><v/></c>`, i.e. a cell with an
*empty* value. Occupancy must therefore be "a `<c>` element exists **or** the cell is inside an
`<f t="array" ref>` range".

**Result of the verified run** (template: table `B2:D5` with a totals row and a calculated column,
`=SEQUENCE(5)` spilling at `F8:F12`, `=SORT(UNIQUE(Tbl[Region]))` spilling at `J3:J4`; edit: insert
3 rows at row 5, write 3 data rows, resize the table, delete `calcChain`, `fullCalcOnLoad`):

```
A_proto.xlsx / A_doc.xlsx          (built by §9's code, run from the file and again from the
  table: $B$2:$D$8 totals: $B$8:$D$8        listing copied out of this document — same result)
   North 10.0 20.0 | South 20.0 40.0 | East 30.0 60.0 | West 40.0 80.0 | North 55.0 110.0
  F11: spill=$F$11:$F$15 f2='=SEQUENCE(5)'      values=[1.0, 2.0, 3.0, 4.0, 5.0]   H11=15.0
  J3 : spill=$J$3:$J$6  f2='=SORT(UNIQUE(Tbl[Region]))'
                                                values=['East','North','South','West']
  parts dropped: xl/calcChain.xml   changed: [Content_Types].xml, workbook rels, table part,
                                            workbook.xml, sheet part
  xl/metadata.xml and xl/styles.xml byte-identical; Excel re-saved the file without repair
```

Both halves of the question are answered:

* the spill that the edit **moved** kept working (`F8:F12` → `F11:F15`, all five values, `=SUM(F11#)`
  still 15);
* the spill whose **size depends on the table** was recalculated by Excel on open and grew from
  `J3:J4` to `J3:J6` — the editor does not need to know the new spill size, only to leave the workbook
  consistent and let `fullCalcOnLoad` do the rest;
* `xl/metadata.xml`, the `cm="1"` index, `t="str"` children and `xl/styles.xml` were untouched
  (byte-identical), and Excel re-saved the edited package without complaint.

What this rules out: an editor cannot silently skip spilled arrays (Rule 2), cannot keep `calcChain`
(Rule 1), and cannot rely on "the cells looked empty" for occupancy (Rule 3).

### 3.6 Proposed module layout and effort

New package `src/fillxl/_ooxml/` (all stdlib: `zipfile`, `xml.etree.ElementTree`):

| Module | Responsibility | est. LOC |
|---|---|---|
| `package.py` | open/copy/write zip, part list, rels graph, content types, byte-preserving save | 350 |
| `xmlio.py` | parse/serialize with the prefix registry; byte-patch helpers for attributes | 150 |
| `workbook.py` | sheets, defined names (global + scoped), `bookViews`, `calcPr`, `calcChain` removal | 270 |
| `worksheet.py` | rows/cells read/write, style refs, insert/delete rows, array/spill `ref` shifting, merges, CF, DV, hyperlinks, dimension | 680 |
| `strings.py` | shared-strings reuse / inline strings | 120 |
| `styles.py` | numFmt lookup + cellXf append (dates only) | 200 |
| `tables.py` | table parts, refs, autoFilter, columns, totals, calculated columns | 300 |
| `charts.py` | chart part discovery, `<f>`/`ptCount` rewriting, multi-level caches | 180 |
| `pivots.py` | `refreshOnLoad`, cache source refs | 100 |
| `sheets_ops.py` | hide/delete/activate + dangling-reference scan in XML | 200 |
| `_backends/_xmledit.py` | `Backend` protocol implementation, wiring to `_plan`/`_charts`/`_validation` | 400 |

≈ **2 850 new LOC**, plus ~400–600 LOC of refactoring in `_cells.py`/`_tables.py`/`_dangling.py`/
`_utils.py` to consume a small "sheet model" instead of openpyxl objects, plus tests.

Rough schedule for one developer familiar with the code base:

| Phase | Work | Duration |
|---|---|---|
| 1 | package I/O, workbook/sheet model, named cells, table fill without resizing, style copy | 1–1.5 weeks |
| 2 | row insert/remove + all reference shifting, table refs, autofilter, totals rows, calculated columns | 1.5–2 weeks |
| 3 | charts, conditional-format extension, pivot refresh, sheet hide/delete/activate, dangling detection | 1–1.5 weeks |
| 4 | Excel-validation harness, parity suite against win32com, docs, deprecation note for the openpyxl engine | 1–1.5 weeks |

**Total ≈ 4–6 weeks** to feature parity on `.xlsx`/`.xlsm`, with the existing test suite as the
safety net.

### 3.7 Risks

| Risk | Mitigation |
|---|---|
| Silently corrupting a template variant nobody tested | keep `_lossy.find_lost_features(template, output) == []` as an assertion; add Excel-open validation to CI (§5) |
| The sheet part *is* rewritten, so unmodelled attributes/elements in it (`cm`/`vm`, `spans`, `extLst`, x14 extensions) could be dropped | model only rows/cells and keep every other element and attribute verbatim; never "clean up" a part |
| Row shifting missed for one dependent feature (drawings, defined names, other tables, array/spill `ref`s) | enumerate them in `_plan`-style unit tests; refuse to grow into non-empty regions (current behaviour); refuse an insert that straddles an array range (§3.5) |
| Stale `xl/calcChain.xml` making Excel reject the file | always delete the part, its `<Override>` and its workbook relationship after any row shift (§3.5 Rule 1) |
| Spilled formula sitting in the growth path looks empty (`<c t="str"><v/></c>`) | treat every `<f t="array" ref>` range as occupied, not just cells with values (§3.5 Rule 3) |
| `mc:Ignorable`/prefix breakage | prefix registry + refuse-unknown-prefix policy (or `lxml`) |
| Two XML code paths (parse vs patch) drift | patch only attributes; everything structural goes through the model |
| New engine is *another* thing to maintain (3 engines) | retire the openpyxl engine for `.xlsx`/`.xlsm` once parity is proven; keep it only as a fallback for one release |

---

## 4. Option B — a backend on excelize-py

### 4.1 What it is

A Python binding of the Go library **excelize** (`v2.11.1`), generated with `c-for-go`:
`main.go` (132 KB) exports `//export`ed C functions, `excelize.py` (298 KB) is a hand-maintained
ctypes wrapper, `types_py.py` holds 67 dataclasses (83 classes) mirroring the Go structs.

| Fact | Value |
|---|---|
| PyPI | `excelize==0.1.0`, single wheel `py3-none-any`, **29.2 MB**, uploaded 2026-07-12 |
| Installed binaries | 8 shared libraries (`libexcelize.{386,amd64,arm64}.{linux.so,darwin.dylib,windows.dll}`) — all platforms installed everywhere (`setup.py`) |
| Platforms | Windows amd64/arm64/i686, macOS amd64/arm64, Linux amd64/arm64/i686 |
| Python | ≥ 3.9 |
| Licence | BSD-3-Clause |
| Formats | XLSX / XLSM / XLAM / XLTX / XLTM (**no** `.xls`, **no** `.xlsb`) |
| Encrypted workbooks | yes — the open options accept a password; **verified**: opened a password-protected `.xlsx` that openpyxl rejects with `BadZipFile` |
| Project health | 51 stars, 31 forks, one maintainer, 4 open issues, `go.mod` pins a *pseudo-version* of excelize (`v2.11.1-0.20261011032807`) |
| Write performance (20k rows × 8 cols, per-row `set_sheet_row`) | **5.90 s** vs openpyxl `append` **3.20 s** (single-threaded, one cgo call per row) |

### 4.2 API inventory against fillxl's needs

Available and sufficient:

| Need | API |
|---|---|
| named cells, global and sheet-scoped | `get_defined_name()`, `set_defined_name(DefinedName(scope=…))` |
| cell writes | `set_cell_value/str/float/int/bool/formula/default`, `set_sheet_row`, `set_cell_rich_text` |
| style copy | `get_cell_style` → `set_cell_style` (verified it applies an existing `s=` index) |
| row insert/delete | `insert_rows`, `remove_row`, `duplicate_row` |
| sheet state | `get_sheet_list`, `get_sheet_visible`, `set_sheet_visible`, `delete_sheet`, `set_active_sheet` |
| recalc on open | `get_calc_props` / `set_calc_props(full_calc_on_load=True)` — verified |
| merge/data-validation readers (overflow checks) | `get_merge_cells`, `get_data_validations` |
| rich text | `get_cell_rich_text` / `set_cell_rich_text` |

Missing or insufficient (the blockers):

| # | Need | Situation |
|---|---|---|
| **B1** | totals rows | `adjustTable` does `t.TableType, t.TotalsRowCount, t.ConnectionID = "", 0, 0` (`adjust.go:692`) on **every** row insert/delete → `totalsRowCount` is erased. Measured: template `ref="B2:D5" totalsRowCount="1"` → after `insert_rows(5, 3)`: `ref="B2:D8"` with `totalsRowCount` gone. Excel then **refused to open the file**. |
| **B2** | calculated columns | `xlsxTableColumn` (`xmlTable.go:179`) has **no** `CalculatedColumnFormula` field, so re-marshalling the table drops `<calculatedColumnFormula>`. Measured: dropped on resize. |
| **B3** | autofilter range | `adjustTable` sets `t.AutoFilter.Ref = t.Ref` — measured result `autoFilter ref="B2:D8"` on a table whose totals row is row 8 (should be `B2:D7`) |
| **B4** | `update_charts` | no `GetCharts`; only `add_chart`, `add_chart_sheet`, `delete_chart`. Measured: after resizing, `SERIES(…,$B$3:$B$4,…)` stays stale. Recreating charts via `add_chart` would discard styling. |
| **B5** | conditional-format extension | `set_conditional_format` / `unset_conditional_format` exist, **no getter** → existing rules cannot be re-emitted with a bigger `sqref` |
| **B6** | pivot refresh | `add_pivot_table` writes `refreshOnLoad="true"` for *new* caches; nothing can set it on an existing cache. Measured: inserting a row into a pivot's source sheet left `worksheetSource ref="A1:B4"` stale and `refreshOnLoad` untouched. |
| **B7** | resizing an existing table | `AddTable` returns `ErrExistsTableName` for a name that already exists; there is no `SetTable`. The only route is the row-insert path, which triggers B1–B3. |
| **B8** | template introspection | `get_tables()` returned `[]` for a table that exists, whenever the sheet relationship used an **absolute** target (`Target="/xl/tables/table1.xml"`, which is what openpyxl writes). With a relative target it worked. Failure is silent. |

### 4.3 Measured behaviour

**Preservation is genuinely good.** On a template with injected modern parts (slicers, timelines,
`chartEx`, `chartStyle`, threaded comments + persons, `richData`, data model, `ctrlProps`, ActiveX,
OLE embeddings, `vbaProject.bin`, `connections`, `queryTables`, `externalLinks`, webextensions,
`customXml`, `metadata.xml`, `calcChain`), an excelize open→write→save round trip kept **every part**;
`find_lost_features()` returned `[]`. Also verified as preserved, byte-semantically:

* `<f t="array" ref="E1:E3">SEQUENCE(3)</f>` (dynamic array) — untouched
* `<f t="shared" si="0" ref="F1:F3">` — untouched
* cell attributes `cm="1" vm="1" ph` (rich-data / data-type metadata) — untouched
* `t="d"` ISO-date cells — untouched
* `cellXfs` indices keep their meaning (font/fill/numFmt mapping identical)
* sheet `state="hidden"`, active sheet, `calcPr`

Also measured, and relevant for design decisions:

* **It rewrites parts it only "knows".** `xl/styles.xml` grew 2 751 → 8 136 bytes, `xl/theme/theme1.xml`
  was re-serialised (entities decoded, whitespace dropped), and all rels parts were re-emitted with
  attributes reordered — even for a no-op save. Content is equivalent, but byte-level golden tests
  would be noisy.
* **A drawing it rewrote was rejected by Excel.** For a template whose drawing part used
  `<c:chart>` with the `c:` prefix bound on an ancestor element, excelize emitted `<c:chart>` without
  re-declaring `c:` → invalid XML → Excel refused the file. Excel-authored drawings (explicit `xdr:`
  prefixes, `c:` bound inline) round-tripped fine, so this is a bug triggered by openpyxl-style
  drawings, in the tested build.
* **Values.** `set_cell_value(date)` writes a serial **with its own style index**, i.e. it does not
  keep the template cell's format; `set_cell_value("=1+1")` writes *text* (good for
  `escape_formulas=True`), `set_cell_formula` is needed for formulas; `set_cell_value(None)` removes
  the cell; `float("nan")` becomes the string `"nan"`.

### 4.4 Verdict for excelize

B1–B6 are not configuration issues; they are holes in the Go library's model that the binding cannot
paper over. The only way to keep fillxl's documented behaviour would be to **post-process the saved
zip** — i.e. re-add `totalsRowCount`, `calculatedColumnFormula`, correct `autoFilter`, rewrite chart
`<f>`, set `refreshOnLoad`. That is exactly the XML editor from §3, added on top of a 29 MB compiled
dependency, with the extra constraint that it must run after someone else's serializer.

Note also that fillxl's `on_lossy` would **not notice any of this**:
`find_lost_features(final_tpl, final_excelize) == []` even though the totals row and the calculated
column were destroyed — the detector compares *parts*, not part contents (§6.2, §6.6).

If fillxl ever wants excelize, the defensible niche is narrow: **reading/opening workbooks openpyxl
cannot** (encrypted files), or bulk data loading where the streaming writer applies.

---

## 5. Test strategy for whichever engine is chosen

* **Keep openpyxl as a *test-only* dependency.** The suite already uses it in ~660 places to author
  fixtures and to verify results independently — that works fine against a backend that does not
  depend on it, and it gives a genuinely independent reader.
* **Excel-open validation.** The box has Excel; there is already an `excel`/`win32com` marker.
  Add a golden-output test that opens every produced file with COM and asserts no repair prompt.
  This is what caught B1/B3 and the drawing bug above.
* **Byte-level part diff.** For the XML engine, assert that untouched parts are byte-identical and
  that the set of rewritten parts is exactly the expected one — cheap, and it is the whole point of
  the design.
* **Invert `on_lossy`.** `find_lost_features(template, output) == []` for the XML engine.
* **Reuse the parity suite** (`tests/parity/`, `tests/integration/test_cross_backend_equivalence.py`)
  with the new engine substituted for openpyxl.

---

## 6. Incidental findings (useful regardless of the decision)

1. **Excel rejects totals-row tables that lack per-column totals metadata.** A table part with
   `totalsRowCount="1"` but no `totalsRowFunction`/`totalsRowLabel` on its columns makes Excel 16
   refuse the file (`Workbooks.Open` fails). Excel's own output always carries them. Verified by
   bisection: adding `totalsRowLabel="Total"` / `totalsRowFunction="sum"` fixed it. Any synthetic
   test template with a totals row needs this, or the Excel-validation harness will report false
   failures.
2. **`_lossy.py` has a blind spot for cell metadata.** openpyxl drops `xl/metadata.xml` **and** the
   `cm`/`vm` cell attributes (measured). `find_lost_features()` returns `[]` for that case, because
   `xl/metadata.xml` is not in `_PART_FEATURES`. This is precisely the "dynamic arrays / data types"
   damage users notice, and the current safety net does not report it. Adding
   `(r"^xl/metadata\.xml$", "cell metadata (dynamic arrays, data types)")` to `_PART_FEATURES` would
   make `on_lossy` honest today. (`xl/pivotCache/…` and `xl/pivotTables/…` are likewise absent from
   the list — worth auditing.)
3. **openpyxl on a template carrying an array formula, a shared formula and cell metadata**
   (measured): the array formula is kept, the shared formula is flattened, `t="d"` is normalised to a
   serial, and `cm`/`vm` plus `xl/metadata.xml` are lost.
4. **excelize's `get_tables` silently returns nothing for absolute relationship targets**, which is
   what openpyxl emits — a template-introspection trap for anyone building on it.
5. **ElementTree prefix rewriting** (measured): without `register_namespace`, an OOXML part round
   trips to `ns0:`/`ns1:` prefixes, which breaks `mc:Ignorable` (a list of *prefixes*, not URIs).
6. **`on_lossy` is a part-level check, not a content-level one.** Measured:
   `find_lost_features(final_tpl, final_excelize) == []` although the table's totals row and its
   calculated column were destroyed in that file. Any new engine therefore needs *content* assertions
   (table `ref`/`totalsRowCount`/`calculatedColumnFormula`, chart series formulas, `sqref`s) in
   addition to the part diff, otherwise silent damage passes the safety net.
7. **`xl/calcChain.xml` is a landmine for any editor that renumbers rows** — not only for fillxl's new
   engine. A stale chain makes Excel refuse the whole workbook rather than recompute (§3.5). Any
   tool that shifts rows must delete it (part + `<Override>` + workbook relationship).
8. **A `t="array"` formula whose `ref` does not match its cell also blocks the file.** `ref` is not
   advisory metadata: Excel validates it on load (§3.5, `A_naive_nochain`).
9. **Excel's own dynamic-array markup is small and stable**: `t="array"` + `ref` + `cm="1"` on the
   master cell, `xl/metadata.xml` with an `XLDAPR` `futureMetadata` entry, cached values on the
   children. Nothing else marks a formula as dynamic — so keeping those parts untouched is sufficient,
   and no special-case code is needed to "understand" `SEQUENCE`/`SORT`/`FILTER`.

---

## 7. Recommendation

1. **Build the direct-OOXML engine** (`engine="xmledit"`, and prefer it from `engine="auto"` for
   `.xlsx`/`.xlsm`). It is the only route that keeps dynamic arrays, slicers, `chartEx`, data types
   and VBA *by construction*, because it stops rewriting what it does not understand.
2. Reuse the parts of fillxl that are already backend-neutral: `_plan.py` (row arithmetic),
   `_charts.update_reference_text` (reference rewriting), `_formulas.resolve_formula`,
   `_validation.py`, `_adapters.py`, `_lossy.py`.
3. Keep `openpyxl` as the fallback engine for one release and as a **test-only** dependency
   afterwards; keep `win32com` for `.xls`/`.xlsb` and format conversion.
4. Fix the `_lossy.py` blind spot in the meantime so `.xlsx` fills stop silently losing cell
   metadata.
5. Do **not** adopt excelize as the primary engine. Revisit only if encrypted workbooks or
   multi-hundred-thousand-row loads become a requirement.

### Files that would change

| File | Change |
|---|---|
| [`src/fillxl/_types.py`](src/fillxl/_types.py) | `Engine` literal gains `"xmledit"` |
| [`src/fillxl/_backend.py`](src/fillxl/_backend.py) | `get_backend`, `resolve_engine`, suffix sets |
| `src/fillxl/_ooxml/**` (new) | the package editor (see §3.6) |
| [`src/fillxl/_backends/_openpyxl.py`](src/fillxl/_backends/_openpyxl.py) | stays, demoted to fallback |
| [`src/fillxl/_cells.py`](src/fillxl/_cells.py), [`_tables.py`](src/fillxl/_tables.py), [`_dangling.py`](src/fillxl/_dangling.py), [`_utils.py`](src/fillxl/_utils.py) | split openpyxl-specific accessors out of the shared logic |
| [`src/fillxl/_lossy.py`](src/fillxl/_lossy.py) | new feature patterns; used as an assertion for the new engine |
| [`docs/backends.md`](docs/backends.md), [`docs/api.md`](docs/api.md), `CHANGELOG.md`, `pyproject.toml` (dev/test extras) | engine matrix, `on_lossy` semantics, dependency move |

---

## 8. Appendix — evidence and reproduction

All experiments ran in throw-away directories (`%TEMP%\excelize_research` for the
openpyxl/excelize comparison, `%TEMP%\spill_research` for the spilled-array round) with fresh venvs;
both were deleted afterwards, so the listings in §9 and the measurements below are the artefact:

```powershell
python -m venv $env:TEMP\excelize_research\venv
& "$env:TEMP\excelize_research\venv\Scripts\python.exe" -m pip install excelize openpyxl pywin32 attrs
python -m venv $env:TEMP\spill_research\venv
& "$env:TEMP\spill_research\venv\Scripts\python.exe" -m pip install pywin32 openpyxl
```

Key measurements, verbatim:

```
# excelize round trip of a rich template
parts dropped  : (none)
parts added    : (none)
parts changed  : [Content_Types].xml, _rels/.rels, xl/_rels/workbook.xml.rels, xl/styles.xml,
                 xl/theme/theme1.xml, xl/workbook.xml, xl/worksheets/sheet1.xml
find_lost_features(template_rich, rt_excelize) -> []

# dynamic arrays / metadata survive excelize
<c r="E1" cm="1" vm="1"><f t="array" ref="E1:E3">SEQUENCE(3)</f><v>1</v></c>
<c r="F1"><f t="shared" ref="F1:F3" si="0">A1*2</f><v>2</v></c>
<c r="E4" t="d" vm="2"><v>2024-01-01</v></c>

# openpyxl loses them
<c r="E1"><f t="array" ref="E1:E3">SEQUENCE(3)</f><v /></c>
<c r="F1"><f>A1*2</f><v /></c>
<c r="E4" t="n"><v>45292</v></c>          # cm/vm gone, xl/metadata.xml gone

# excelize table resize on a template with a totals row and a calculated column
before: <table ... ref="B2:D5" headerRowCount="1" totalsRowCount="1">
          <autoFilter ref="B2:D4"/><tableColumns count="3">
            <tableColumn id="1" name="Region" totalsRowLabel="Total"/>
            <tableColumn id="2" name="Amount" totalsRowFunction="sum"/>
            <tableColumn id="3" name="Double"><calculatedColumnFormula>C3*2</calculatedColumnFormula></tableColumn>
after : <table ... ref="B2:D8" headerRowCount="1"><autoFilter ref="B2:D8"/>
            <tableColumn id="1" name="Region" totalsRowLabel="Total"/>
            <tableColumn id="2" name="Amount" totalsRowFunction="sum"/>
            <tableColumn id="3" name="Double"/>            <-- calculated formula lost, totalsRowCount lost
chart : SERIES(Data!$C$2,Data!$B$3:$B$4,Data!$C$3:$C$4,1)  <-- stale
Excel : FAILED (file refused)

# direct-XML prototype on the same template
table : ref="B2:D8" totalsRowCount="1" <autoFilter ref="B2:D7"/>
        <calculatedColumnFormula>C3*2</calculatedColumnFormula> kept
chart : SERIES(Data!$C$2,Data!$B$3:$B$7,Data!$C$3:$C$7,1)
parts : only 4 rewritten (sheet, table, chart, workbook); every other part byte-identical
Excel : OK  table=$B$2:$D$8 totals=$B$8:$D$8  Double D3==C3*2  totals=155
find_lost_features(template, prototype) -> []

# Excel accepts an excelize round trip of an Excel-authored chart workbook
xl_chart.xlsx     OPEN OK  C3=10.0
xl_chart_rt.xlsx  OPEN OK  C3=77.0     (drawing rewritten with explicit xdr: prefixes)

# encrypted workbook
excelize.open_file("encrypted.xlsx", options with the workbook password) -> OK, A1 = "secret"
openpyxl.load_workbook(encrypted.xlsx) -> BadZipFile: File is not a zip file

# inline strings next to an existing shared-strings table (Excel-authored file)
xl_mixed_strings.xlsx  OPEN OK  B7 = "inline-added"  B3 = "North"   (B3 comes from xl/sharedStrings.xml)

# write throughput, 20 000 rows x 8 columns
excelize set_sheet_row : 5.90 s
openpyxl append        : 3.20 s

# ── spilled arrays (§3.5): template authored by Excel, rows inserted by the prototype ──
# template A.xlsx: table B2:D5 (totals row + calculated column), F8=SEQUENCE(5) spilling F8:F12,
#                  H8=SUM(F8#), J3=SORT(UNIQUE(Tbl[Region])) spilling J3:J4, xl/metadata.xml present
A.xlsx                 OPEN OK  F8 spill=$F$8:$F$12  values=[1,2,3,4,5]  H8=15  J3 spill=$J$3:$J$4
A_naive.xlsx           OPEN FAILED: Microsoft Excel     (rows shifted, array ref stale, calcChain kept)
A_naive_nochain.xlsx   OPEN FAILED: Microsoft Excel     (same, calcChain removed  -> stale ref is fatal)
A_refs.xlsx            OPEN FAILED: Microsoft Excel     (ref shifted correctly, calcChain kept -> fatal)
A_refs_nochain.xlsx    OPEN OK   F11 spill=$F$11:$F$15  values=[1,2,3,4,5]  H11=15 (SUM(F11#))
A_full_nochain.xlsx    OPEN OK   table=$B$2:$D$8 totals=$B$8:$D$8
                                Double column = [20,40,60,80,110]  totals row = [0,31,5]
                                F11 spill=$F$11:$F$15 values=[1,2,3,4,5]
                                J3  spill=$J$3:$J$6  values=['East','North','South','West']  (grew from J3:J4)
# blank spill children are written as empty cells, so value-based occupancy checks can miss them:
<c r="C8" t="str"><v/></c>

# the prototype published in §9, run on A.xlsx -> A_proto.xlsx
A_proto.xlsx           OPEN OK   table=$B$2:$D$8 totals=$B$8:$D$8
                                F11 spill=$F$11:$F$15 f2='=SEQUENCE(5)' values=[1,2,3,4,5] H11=15.0
                                J3  spill=$J$3:$J$6  f2='=SORT(UNIQUE(Tbl[Region]))'
                                values=['East','North','South','West']
                                xl/metadata.xml, xl/styles.xml untouched; Excel re-saved it without repair
```

Excel-side checks used COM (`Excel.Application`, `DisplayAlerts=False`, `AutomationSecurity=3`,
`Workbooks.Open(path, 0, True)`) and read `ListObjects(...).Range`, `.TotalsRowRange`,
`ListColumns("Double").DataBodyRange.Formula` and `ChartObjects(1).Chart.SeriesCollection(1).Formula`.
Spill state was read with `Range.HasSpill`, `Range.SpillingToRange`, `Range.Formula2` after
`Application.CalculateFullRebuild()`.

---

## 9. Appendix — the OOXML row-shift prototype (runnable)

This is the complete code that produced `A_proto.xlsx` above: no openpyxl, no Excel, stdlib only.
It grows an Excel table by *n* rows over a template that contains a totals row, a calculated column
and two spilled dynamic arrays.

```python
"""Minimal OOXML fill prototype: grow an Excel table, no openpyxl, no Excel.

Handles the three things a row-shifting edit must get right for dynamic arrays:
  * shift <f t="array" ref="…"> together with the cell that owns it
  * refuse to insert rows strictly inside an array/spill range
  * delete xl/calcChain.xml (a stale chain makes Excel refuse the workbook)
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

CELL_REF = re.compile(r"(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")
ARRAY_REF = re.compile(r'<f[^>]*\bt="array"[^>]*\bref="([^"]+)"')


def split_ref(ref: str) -> tuple[str, int, str, int]:
    """``"B3:D5"`` -> ``("B", 3, "D", 5)``; a single cell becomes a 1x1 range."""
    first, _, last = ref.partition(":")
    last = last or first
    column = lambda c: re.match(r"(\$?[A-Z]+)", c).group(1)      # noqa: E731
    row = lambda c: int(re.search(r"(\d+)$", c).group(1))        # noqa: E731
    return column(first), row(first), column(last), row(last)


def refuse_if_straddling(sheet: str, insert_at: int) -> None:
    """A row insert inside an array range cannot be expressed: refuse it."""
    for ref in ARRAY_REF.findall(sheet):
        _, first_row, _, last_row = split_ref(ref)
        if first_row < insert_at <= last_row:
            msg = f"array/spill range {ref} straddles row {insert_at}; refusing to insert"
            raise ValueError(msg)


def shift_formula(text: str, insert_at: int, count: int) -> str:
    """Move the relative row numbers of a formula that lives on a shifted row."""

    def repl(m: re.Match[str]) -> str:
        if m.group(3) == "$":                                    # absolute row
            return m.group(0)
        row = int(m.group(4))
        return m.group(0) if row < insert_at else f"{m.group(1)}{m.group(2)}{m.group(3)}{row + count}"

    return CELL_REF.sub(repl, text)


def shift_range(text: str, insert_at: int, count: int) -> str:
    """Move an array/spill range that lies entirely below the insert point."""
    c1, r1, c2, r2 = split_ref(text)
    return text if r2 < insert_at else f"{c1}{r1 + count}:{c2}{r2 + count}"


def shift_rows(sheet: str, insert_at: int, count: int, added: dict[int, str]) -> str:
    """Renumber rows at/after *insert_at* and splice *added* rows in."""
    rows = {int(re.match(r'<row r="(\d+)"', x).group(1)): x
            for x in re.findall(r"<row [^>]*>.*?</row>", sheet, re.S)}
    moved: dict[int, str] = {}
    for number, xml in rows.items():
        target = number + count if number >= insert_at else number
        if target != number:
            xml = re.sub(rf'(<row r="){number}(")', rf"\g<1>{target}\g<2>", xml)
            xml = re.sub(rf'(<c r="(?:\$?[A-Z]+\$?)){number}(")', rf"\g<1>{target}\g<2>", xml)
            xml = re.sub(
                r"<f([^>]*)>([^<]*)</f>",
                lambda m: "<f{}>{}</f>".format(
                    re.sub(r'(ref=")([^"]+)(")',
                           lambda a: f'{a.group(1)}{shift_range(a.group(2), insert_at, count)}{a.group(3)}',
                           m.group(1)),
                    shift_formula(m.group(2), insert_at, count),
                ),
                xml,
            )
        moved[target] = xml
    moved.update(added)
    return "<sheetData>" + "".join(moved[n] for n in sorted(moved)) + "</sheetData>"


def data_row(row: int, values: tuple[object, ...], first_col: str = "B") -> str:
    """One table body row: text cells inline, numbers as numbers, formulas as <f>."""
    letters = [chr(ord(first_col) + i) for i in range(len(values))]
    cells = ""
    for letter, value in zip(letters, values):
        ref = f"{letter}{row}"
        if isinstance(value, str) and value.startswith("="):
            cells += f'<c r="{ref}"><f>{value[1:]}</f></c>'
        elif isinstance(value, str):
            cells += f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'
        else:
            cells += f'<c r="{ref}"><v>{value}</v></c>'
    return f'<row r="{row}">{cells}</row>'


def strip_calc_chain(parts: dict[str, bytes]) -> None:
    """Drop calcChain: it goes stale on any row shift and then blocks the file."""
    parts.pop("xl/calcChain.xml", None)
    parts["[Content_Types].xml"] = re.sub(
        rb'<Override PartName="/xl/calcChain\.xml"[^>]*/>', b"", parts["[Content_Types].xml"])
    parts["xl/_rels/workbook.xml.rels"] = re.sub(
        rb'<Relationship [^>]*calcChain\.xml"[^>]*/>', b"", parts["xl/_rels/workbook.xml.rels"])


def fill(template: Path, output: Path, rows: list[tuple[object, ...]], *, table_ref: str,
         new_table_ref: str, new_autofilter_ref: str, insert_at: int) -> None:
    """Insert ``len(rows)`` rows at *insert_at* and resize the table part."""
    with zipfile.ZipFile(template) as z:
        parts = {name: z.read(name) for name in z.namelist()}

    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    refuse_if_straddling(sheet, insert_at)
    added = {insert_at + i: data_row(insert_at + i, row) for i, row in enumerate(rows)}
    sheet = re.sub(r"<sheetData>.*?</sheetData>",
                   lambda _: shift_rows(sheet, insert_at, len(rows), added),
                   sheet, flags=re.S)
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()

    table = parts["xl/tables/table1.xml"].decode()
    table = table.replace(f'ref="{table_ref}"', f'ref="{new_table_ref}"', 1)
    table = re.sub(r'(<autoFilter[^>]*\bref=")[^"]*"', rf'\g<1>{new_autofilter_ref}"', table)
    parts["xl/tables/table1.xml"] = table.encode()

    parts["xl/workbook.xml"] = re.sub(
        rb"<calcPr[^>]*>", b'<calcPr calcId="191029" fullCalcOnLoad="1"/>', parts["xl/workbook.xml"])

    strip_calc_chain(parts)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in parts.items():
            z.writestr(name, blob)


if __name__ == "__main__":
    template, output = Path(sys.argv[1]), Path(sys.argv[2])
    fill(
        template,
        output,
        [("East", 30, "=Tbl[[#This Row],[Amount]]*2"),
         ("West", 40, "=Tbl[[#This Row],[Amount]]*2"),
         ("North", 55, "=Tbl[[#This Row],[Amount]]*2")],
        table_ref="B2:D5",           # header + 2 body rows + totals row
        new_table_ref="B2:D8",
        new_autofilter_ref="B2:D7",  # autofilter excludes the totals row
        insert_at=5,                 # first row below the table body
    )
    print("wrote", output)
```

Run it as `python ooxml_fill.py template.xlsx output.xlsx`. What it deliberately does **not** do:
this is a feasibility prototype, not the engine — it hard-codes the sheet part and the table part,
does not shift `<mergeCell>`, `<conditionalFormatting>`, `<dataValidation>`, `<hyperlink>`, drawing
anchors, other tables, defined names or charts, and does not model styles. The production design
(§3.6) covers those; the point of this listing is that the spill-specific parts of the problem —
`ref` shifting, straddle refusal, `calcChain` removal — are small and testable.

```powershell
# how the published listing was validated
& "$env:TEMP\spill_research\venv\Scripts\python.exe" ooxml_fill.py A.xlsx A_proto.xlsx
# then opened in Excel 16: table $B$2:$D$8, totals $B$8:$D$8,
# F11 spill $F$11:$F$15 = [1,2,3,4,5], H11 = 15, J3 spill $J$3:$J$6 = ['East','North','South','West']
```
