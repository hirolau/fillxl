# Benchmarks

Filling is dominated by the engine, so the useful comparison is between
engines **in the same run**. The numbers below come from
[`scripts/benchmark.py`](https://github.com/hirolau/fillxl/blob/main/scripts/benchmark.py)
and are published as a reference point, not as a promise: absolute times follow
the CPU, the disk and the size of the template.

!!! note "Measure your own machine"
    ```bash
    FILLXL_BENCH_ROWS=2000 FILLXL_BENCH_ITER=3 python scripts/benchmark.py
    ```
    `FILLXL_BENCH_ROWS` and `FILLXL_BENCH_ITER` default to 5,000 rows and 10
    iterations. Excel-only methods are skipped automatically when Excel is not
    installed.

## How it is measured

- Template: `tests/assets/test_template_5_wide_table.xlsx`, table `CityData`,
  20 columns. Its 10 rows are repeated to the requested row count.
- One warm-up call per method, then 3 timed iterations; the table shows the
  mean ± standard deviation.
- Timed region: the whole `fill()` call — opening the template, filling it and
  writing the output to a temporary directory. Reading the data from disk is
  not included.
- Machine: AMD Ryzen 7 7435HS, 16 threads, Windows 11, Python 3.12,
  openpyxl 3.1.5, Microsoft Excel 16.

## Engines compared on the same fill

2,000 rows × 20 columns, seconds, lower is better:

| Data source | `openpyxl` | `fillxl` | `win32com` |
|---|---:|---:|---:|
| `list[dict]` | 1.49 ± 0.03 | **0.70 ± 0.01** | 2.59 ± 0.10 |
| polars `DataFrame` | 1.97 ± 0.12 | **0.89 ± 0.01** | 2.43 ± 0.02 |
| pandas `DataFrame` | 1.95 ± 0.03 | **0.93 ± 0.01** | 2.54 ± 0.04 |
| duckdb relation | 1.54 ± 0.04 | **0.71 ± 0.01** | 2.58 ± 0.01 |

For scale, two methods that write the same data without a template:

| Method | Seconds |
|---|---:|
| `pandas.to_excel` (openpyxl) | 1.66 ± 0.08 |
| `pandas.to_csv` | 0.03 ± 0.00 |

`engine="fillxl"` comes out roughly **twice as fast as openpyxl** for the same
workbook. It writes the same values but never builds a workbook object model or
re-serialises parts it did not change. `win32com` pays for starting Excel and
for the COM round trips, and it is the only engine whose time includes
recalculating the workbook.

## Fill time grows with the row count, not with its square

`engine="fillxl"`, a 5-column table sharing its sheet with a spilled array,
one iteration per run:

| Rows | Seconds |
|---:|---:|
| 250 | 0.14 |
| 1 000 | 0.12 |
| 2 000 | 0.22 |
| 4 000 | 0.37 |

Sixteen times the rows cost under three times the time — the fixed cost of
opening and writing the package dominates at these sizes, and the per-row work
is a constant amount of text editing.

## What these numbers do not tell you

- **Cached values.** Only `win32com` computes formulas; `openpyxl` and `fillxl`
  leave the workbook flagged `fullCalcOnLoad`, so Excel spends its own time on
  first open. See [The fillxl engine](fillxl-engine.md#cached-values).
- **Your template.** A workbook with many sheets, charts and pivot caches
  rewrites more parts; a narrow table writes fewer cells. Both change the
  ratio between the engines.
- **Your data.** Data already in memory is assumed. Converting a 500 MB CSV
  will cost far more than any of the fills above.
