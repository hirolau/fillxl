# fillxl engine — implementation log

Goal: a fillxl backend that **does not need Microsoft Excel** and **keeps spilled
(dynamic-array) ranges** and every other feature the openpyxl engine drops.

The research behind the design is in [BACKEND-RESEARCH.md](BACKEND-RESEARCH.md).

---

## Decisions

### D1 — The engine edits the OOXML package directly

`engine="fillxl"` reads the `.xlsx`/`.xlsm` ZIP, copies every part byte for byte and
rewrites only the parts that must change (worksheet XML, table parts, chart parts,
pivot cache definitions, `xl/workbook.xml`, `[Content_Types].xml`, one `.rels` file).

Why: openpyxl's damage is not that it reads XML badly — it rebuilds the whole package
from its object model, so anything the model cannot represent (spill metadata
`xl/metadata.xml`, `cm`/`vm` cell attributes, slicers, `chartEx`, rich data, VBA)
disappears. A surgical editor cannot drop what it never parses.

### D2 — Parts are edited as text, not re-serialised

`xml.etree.ElementTree` renames namespace prefixes and drops namespace declarations
that only `mc:Ignorable` refers to, which makes Excel refuse the workbook. Only the
pieces that change are rewritten; everything else in a part keeps its exact bytes.
(Reading — never writing — with ElementTree is fine and is used for `.rels` files.)

### D3 — The engine never inserts or removes worksheet rows

The research prototype inserted rows, which forced it to shift merges, data
validations, hyperlinks, drawing anchors, defined names, chart references and
array-formula `ref`s, and to delete `xl/calcChain.xml`.

Every fillxl engine already refuses to grow a table into occupied cells
(`_check_overflow`), so growth only ever lands in *verified empty* cells. Adding
cells instead of shifting rows therefore keeps **all** of that content valid by
construction: spills, merges, drawings and `calcChain` never move. This is the same
contract the openpyxl and win32com engines implement, so parity is preserved.

Consequence: the "row shifting" half of the research design is not needed, and the
two hard rules that came out of it (delete a stale `calcChain`, shift
`<f t="array" ref>`) cannot be violated. The remaining rule — **an array/spill range
is occupied even when its cells look empty** — is implemented, because a spill
reserves cells without storing values in them.

### D4 — `xl/calcChain.xml` is dropped on save

Not because we move rows (we do not) but because a chain that disagrees with the
workbook makes Excel refuse the file, and Excel rebuilds the chain on load anyway.
Cheap insurance; verified to change nothing else in the output.

### D5 — `fullCalcOnLoad="1"`

The engine writes formulas without cached results (like the openpyxl engine), so
Excel must recalculate on open.

### D6 — Text is written as inline strings

Appending to `xl/sharedStrings.xml` means maintaining `count`/`uniqueCount` and
never renumbering existing entries — a whole class of corruption for no benefit.
`t="inlineStr"` leaves `sharedStrings.xml` byte-identical.

### D7 — Scope limits (documented, enforced with clear errors)

* `.xlsx` / `.xlsm` only; `.xlsb`, `.xls` and format conversion stay on `win32com`.
* Password protected (CFB/ECMA-376 agile) workbooks are refused, not misread.
* Modern charts (`chartEx`) are left untouched — Excel does not express their sources
  as the `<c:f>` references this engine rewrites, and touching them is worse than
  leaving them alone.
* `openpyxl` remains a *library* dependency, but the engine only uses its pure-Python
  helpers (`CellRange`, `get_column_letter`, formula `Tokenizer`, number-format
  predicates) — none of which need Excel or parse a workbook.

---

## Task log

### Phase 0 — scaffolding
- [x] `Engine` literal gains `"fillxl"`; `resolve_engine`/`get_backend` wired
- [x] Feature branch `feature/fillxl-native-engine`

### Phase 1 — package primitives (`src/fillxl/_ooxml/`)
- [x] `xmltext.py` — escape/unescape, element regexes, attribute surgery
- [x] `package.py` — ZIP parts, byte-preserving save, content types, relationships
- [x] `workbook.py` — sheets, defined names, `calcPr`, active tab, sheet state, sheet deletion
- [x] `worksheet.py` — rows/cells, styles, occupancy, array ranges, conditional formatting, formulas
- [x] `cells.py` — value → `<c>` XML (inline text, numbers, booleans, dates, formulas)
- [x] `styles.py` — number-format lookup and date `cellXfs` creation
- [x] `tables.py` — table parts: `ref`, `autoFilter`, columns, calculated columns
- [x] `charts.py` — classic chart `<c:f>` and cache `ptCount` rewriting
- [x] `pivots.py` — `refreshOnLoad`

### Phase 2 — fill semantics
- [x] named cells (global + sheet-scoped, case-insensitive fallback, header guard)
- [x] table fill: plan-based growth/shrink, totals row move, calculated columns, style copy
- [x] conditional-formatting extension
- [x] adjacent formula resizing (`_formulas.extend_table_ranges`)
- [x] spill/array refusal for writes inside an array range

### Phase 3 — backend wiring
- [x] `_backends/_fillxl.py`
- [x] hide / delete / activate sheets, dangling-reference scan in XML
- [x] `check_output` (`on_lossy`) for the engine
- [x] `_lossy` blind spot for `xl/metadata.xml` fixed

### Phase 4 — tests & docs
- [x] unit tests for every `_ooxml` module (`tests/unit/test_ooxml.py`)
- [x] integration tests: spills survive a fill (no Excel needed, ZIP + openpyxl read-back)
- [x] Excel COM validation test: produced files open without repair, spills alive
- [x] parity: `"fillxl"` added to the shared cross-backend `ENGINES` list
- [x] docs (`README`, `docs/backends.md`, `docs/api.md`, `CHANGELOG`)

---

## Progress notes

### Bug found by the new tests: attribute on a closing tag

`set_attribute()` works on a start tag. Passing a whole element (`<c>…</c>`,
`<conditionalFormatting>…</conditionalFormatting>`) appended the attribute to
the **closing** tag, producing XML that no reader can parse. It was caught by
adding `assert_well_formed()` to every fillxl test output — the corrupted file
opened fine in no reader, so only an explicit parse assertion surfaces it.
`xmltext.set_tag_attribute()` now makes the mistake impossible.

### Deliberate deviations from the research prototype

* **No row insertion.** The prototype inserted rows and therefore had to shift
  merges, validations, drawings, names and array `ref`s and delete
  `calcChain`. fillxl already refuses to grow into occupied cells, so growth
  lands in verified-empty cells and nothing below the table moves. The two
  hard rules the research derived (delete a stale `calcChain`, shift array
  `ref`s) cannot be violated by construction; the third (array ranges are
  occupied) is implemented.
* **`xl/calcChain.xml` is still dropped** on save — not for row moves, but
  because it costs nothing and removes a whole class of "Excel refuses the
  file" failures.
* **`engine="auto"` still selects openpyxl.** Switching the default is a
  separate, user-visible decision; `engine="fillxl"` is opt-in for this
  release.

### Verification

| Check | Result |
|---|---|
| `pytest tests/ -q` (full suite, Excel available) | 1128 passed, 1 xfailed |
| `pytest --doctest-modules src/fillxl/` | 87 passed |
| `ruff check src/fillxl tests` | clean |
| `mypy src/fillxl` (strict) | clean |
| `pyrefly check` | clean |
| `ty check src/fillxl tests` | 12 pre-existing errors in `tests/unit/test_formulas.py` (not touched by this work) |
| Real Excel 16: spill authored by Excel survives a fill and grows with the table | `HasSpill=True`, `B2:B7`, values 1…6 |
| Real Excel 16: the repository's macro template opens without repair, VBA byte-identical | verified |
