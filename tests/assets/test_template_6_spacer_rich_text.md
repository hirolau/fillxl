# Excel Test File: test_template_6_spacer_rich_text.xlsx — Sheet Descriptions

A workbook written by Excel itself, kept as a golden template for the rich-text
whitespace fix. Every header cell of the `Portfolio Holding` sheet is rich text
built from three runs:

```
" \r\n"                      — a 2.95pt Calibri spacer run
"Instrument Type"            — the 9pt bold white label
"\r\n "                      — the closing spacer run
```

The spacer runs hold nothing but whitespace, so the two `<t>` elements they
produce need `xml:space="preserve"`. Excel writes that attribute itself, but
openpyxl only writes it when the text also contains a visible character, which
made every openpyxl save of this template a file Excel refuses to open
("We found a problem with some content…", `Workbooks.Open` with
`CorruptLoad=0` fails).

Sanitized copy of a real report: all fund, issuer, security, ISIN and vendor-ID
text was replaced with synthetic values, and the company logo was removed. The
structure, styles, number formats and rich-text runs are those Excel produced.

---

## Sheet: Portfolio Holding

`A1:R295`. Rows hold a fund header (title, fund name, disclaimer line, asset
class, AUM, benchmark) and a 282-row holdings table.

| Row | Contents |
|-----|----------|
| 2   | Rich-text title `Portfolio Holding \| <date>` in merged `C2:I2` |
| 4   | Fund name, bold 11pt, merged `A4:I4` |
| 6   | Two-run disclaimer line, merged `A6:I6` |
| 8–10| Asset class, fund AUM and benchmark, merged `A8:I10` |
| 13  | Column headers — 18 rich-text cells with whitespace-only spacer runs |
| 14–295 | Holdings: type, sub type, security name, ISIN, other ID, currency, quantity, prices, values, weight, exposure |
| 296 | Empty hidden row |

## Sheet: Disclaimer

`A1:C3`. One multi-run rich-text cell in `B2` (bold country names inside a body
of 9pt text), plus empty styled border cells.

## Used by

- `tests/integration/test_rich_text.py` — the whitespace-only runs are still
  marked `xml:space="preserve"` after a fill
- `tests/ooxml/test_excel_opens.py` — the filled workbook opens in real Excel
  without a repair prompt
