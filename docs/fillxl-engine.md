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

- `.xlsx` and `.xlsm` only, and only in the same format: no `.xlsb`,
  no encrypted workbooks, no format conversion — those stay on
  `engine="win32com"`.
- Modern charts (`chartEx`) follow a table resize: the `<cx:f>` source of each
  dimension and the cached `ptCount` of its levels are rewritten.  Sparklines
  follow too, on whichever sheet they sit on.
- Pivot tables are flagged for refresh on open; fillxl never pivots.
- Pivot sources, defined names and data validations that cover a table body
  follow a resize, as they do in Excel. A built-in name (`_xlnm.Print_Area`,
  …) is Excel's own and is reported with `UntrackedRangeWarning` instead of
  being guessed at.

## Cached values

The engine does not calculate. Formulas it writes have no cached result, and a
fill removes every cached result already in the workbook: the numbers the
template cached describe the template's data, and a reader that cannot
recalculate would show them as the filled report's. Spills survive, but their
stored results are cleared the same way. Without a dependency graph there is no
telling which cached result is still valid, so all of them go; a sheet with no
formulas is still copied byte for byte. Excel recalculates the file when it
opens it — the workbook is flagged `fullCalcOnLoad` — and a spill whose size
depends on the table grows then.

Tools that read cached values (`pandas`, `openpyxl` with `data_only=True`, file
previews, search indexers) see blanks in a fillxl output. A blank is the
deliberate answer — a stale number is indistinguishable from a correct one.
Chart *data* caches (`numCache`/`strCache` points) are the exception: they keep
the template's points until Excel redraws the chart on open. Dynamic arrays
survive on Linux, but they are not computed there either. Use
`engine="win32com"` when the saved file must already contain computed values.

## When to use it

Reach for `engine="fillxl"` when the template contains something openpyxl
cannot write — a spilled `SEQUENCE`/`FILTER`/`SORT`, slicers, data types,
`chartEx`, VBA — and you still want a fill that runs without Excel, on any
platform.
