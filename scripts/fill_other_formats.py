"""Try filling various Excel file formats with both backends.

Usage:
    uv run python scripts/fill_other_formats.py
"""

from __future__ import annotations

from pathlib import Path

from fillxl import fill

ASSETS = Path(__file__).resolve().parent.parent / "tests" / "assets"
OUTPUT = Path(__file__).resolve().parent.parent / "filled_templates"
TEMPLATE_1 = ASSETS / "test_template_1.xlsx"


def try_fill(label: str, template: Path, engine: str) -> str:
    outdir = OUTPUT / engine / "other_formats"
    outdir.mkdir(parents=True, exist_ok=True)
    output = outdir / f"{template.stem}_filled.xlsx"
    try:
        fill(
            template=template,
            output=output,
            values={"single_cell": "Filled by fillxl"},
            engine=engine,
        )
        return f"  {label:35s} {engine:10s} OK  ({output.stat().st_size:,} bytes)"
    except Exception as e:
        msg = str(e).replace("\n", " | ")
        return f"  {label:35s} {engine:10s} FAIL: {msg[:140]}"


def main() -> None:
    # Create dummy files with various extensions for testing
    dummy_dir = OUTPUT / "other_formats" / "_dummy"
    dummy_dir.mkdir(parents=True, exist_ok=True)
    dummy_files = {}
    for ext in [".xls", ".xltx", ".xltm", ".csv", ".txt", ".zip", ".ods", ".pdf"]:
        p = dummy_dir / f"dummy{ext}"
        if not p.exists():
            p.write_text("not a real excel file")
        dummy_files[ext] = p

    templates = [
        (".xlsx (reference)", TEMPLATE_1),
        (".xlsm", ASSETS / "test_template_1_slsm.xlsm"),
        (".xlsb", ASSETS / "test_template_1_xlsb.xlsb"),
    ]
    for ext, p in dummy_files.items():
        templates.append((ext, p))

    engines = ["openpyxl", "win32com"]

    print(f"\n{'-' * 90}")
    print("  fill_other_formats - trying various extensions with both backends")
    print(f"{'-' * 90}")

    for label, path in templates:
        print(f"\n  File: {label}")
        for engine in engines:
            result = try_fill(label, path, engine)
            print(result)

    print()


if __name__ == "__main__":
    main()
