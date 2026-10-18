# Backends

fillxl supports two backend engines for writing Excel files. Choose based on your platform and fidelity requirements.

Both engines are expected to produce **the same values** for the same input;
the test suite enforces that with a cross-backend equivalence suite.

## Comparison

| Engine | Cross-platform | Requires Excel | Best for |
|--------|:---:|:---:|---|
| `openpyxl` | :material-check: | :material-close: | Most use cases, Linux/macOS |
| `win32com` | :material-close: Windows only | :material-check: | `.xlsm`/`.xlsb`, modern charts, full fidelity |

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

### Limitations

- Does **not** support `.xlsm`, `.xlsb`, or macro-enabled formats
- Drops some modern chart types (e.g., sunburst, treemap)
- No dynamic array formula support

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

### Limitations

- Windows only
- Requires a licensed Excel installation
- Slower per-call overhead (Excel COM startup)

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

Strings starting with `=` become formulas in both engines. Use
`FillXLConfig(escape_formulas=True)` when the data may come from an untrusted
source and a leading `=` must stay literal.

Excel tables always end up with **at least one body row**: filling zero rows
clears and keeps a single row, because Excel rejects a table without a body.

Pivot caches are refreshed by both engines: openpyxl marks them
(`refreshOnLoad`) so Excel recalculates on open, and win32com recalculates them
during the fill as well.

Tables with a **totals row** are rejected with a clear error — the totals row
would be overwritten by the data and the row arithmetic would be off by one.
Remove the totals row from the template (or compute the totals yourself) before
filling.

!!! warning "CI/CD"
    win32com tests are excluded from CI (Linux runners can't run Excel COM).
    This backend is tested manually on Windows before each release.
