# The fillxl engine

`engine="fillxl"` fills a workbook by **editing the OOXML package itself**.
There is no workbook object model in between: no Excel, no `Worksheet` objects,
nothing that can quietly forget a feature on the way out.

```python
from fillxl import fill

fill(
    template="tpl.xlsx",
    output="out.xlsx",
    tables={"Orders": rows},
    values={"report_title": "Q3"},
    engine="fillxl",
)
```

It is **experimental and opt-in**. `engine="auto"` never selects it, so a fill
only uses it when you ask for it. Spilled arrays survive the fill untouched —
Excel computes them when it opens the file (see [Cached
values](#cached-values)). [Backends](backends.md) compares it with the other two
engines and lists its limits.

## What it does

An `.xlsx`/`.xlsm` file is a ZIP of XML parts. The engine:

1. opens the template as a ZIP and keeps every entry's original `ZipInfo`;
2. reads a part only when a fill needs it, and rewrites only the parts the
   fill has to change — the worksheet, the table definition, chart references,
   pivot cache flags and `xl/workbook.xml`;
3. writes every other part back **byte for byte**, streamed straight from the
   template, so a 20 MB image costs nothing to keep.

A part the engine never parses is a part it cannot damage. That is why spilled
dynamic arrays and their `xl/metadata.xml`, `cm`/`vm` cell metadata, slicers,
timelines, `chartEx`, rich data, embedded objects and VBA survive: they are
copied, not rebuilt.

## How it stays safe

- **Rows are never inserted or removed.** A growing table writes into cells
  verified empty — the same contract the other engines implement — so merges,
  data validations, hyperlinks, drawing anchors, defined names, other tables
  and array-formula ranges never move and never go stale.
- **An array range counts as occupied.** A spill can hold blank results that
  look like empty cells; growing a table into one is refused with an error that
  names the array formula.
- **Shared formulas are edited as groups.** Excel stores a filled-down formula
  once, with empty dependents. A group the fill touches is expanded to one
  formula per cell first, so no dependent is left without its master; a group
  outside the fill keeps its exact bytes.
- **A save refuses to write broken XML.** Every part the engine rewrote is
  parsed before the output file is opened; a malformed part raises
  `RuntimeError` asking for a bug report instead of producing a workbook no
  reader can open.
- **`xl/calcChain.xml` is dropped** on save. Excel rebuilds it, and a chain that
  disagrees with the workbook makes Excel refuse the file.
- **Dates follow the workbook.** A `date1904` workbook gets 1904-based serials,
  so dates do not jump four years on open.

## What it does not do

- `.xlsx` and `.xlsm` only, and only in the same format: no `.xlsb`, no `.xls`,
  no encrypted workbooks, no format conversion — those stay on
  `engine="win32com"`.
- Modern charts (`chartEx`) are left untouched, so their ranges do not follow a
  table resize.
- Pivot tables are flagged for refresh on open; fillxl never pivots.
- Pivot sources, defined names and data validations that cover a table body do
  not follow a resize; the fill warns with `UntrackedRangeWarning` instead of
  letting the report miss rows.

## Cached values

The engine does not calculate. Formulas it writes have no cached result, and the
sheet a table fill changes keeps none either: the numbers the template cached
there describe the template's data, and a reader that cannot recalculate would
show them as the filled report's. Spills survive, but their stored results are
cleared the same way. Every other part is copied byte for byte, so an untouched
sheet keeps the template's cached numbers until Excel opens and recalculates the
file; the workbook is flagged `fullCalcOnLoad` so that happens on open. A spill
whose size depends on the table therefore grows only when Excel recalculates.

Tools that read cached values (`pandas`, `openpyxl` with `data_only=True`, file
previews, search indexers) see blanks or old numbers in a fillxl output. A blank
is the deliberate answer — a stale number is indistinguishable from a correct
one. Dynamic arrays survive on Linux, but they are not computed there either.
Use `engine="win32com"` when the saved file must already contain computed
values.

## When to use it

Reach for `engine="fillxl"` when the template contains something openpyxl
cannot write — a spilled `SEQUENCE`/`FILTER`/`SORT`, slicers, data types,
`chartEx`, VBA — and you still want a fill that runs without Excel, on any
platform.
