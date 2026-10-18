# Backends

fillxl supports two backend engines for writing Excel files. Choose based on your platform and fidelity requirements.

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

!!! warning "CI/CD"
    win32com tests are excluded from CI (Linux runners can't run Excel COM).
    This backend is tested manually on Windows before each release.
