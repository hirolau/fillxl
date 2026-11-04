"""Surgical editing of an OOXML workbook package.

This is the machinery behind ``engine="fillxl"``: a `.xlsx`/`.xlsm` file is a
ZIP of XML parts, and a fill only has to change a few of them.  Everything else
is copied byte for byte, which is what keeps spilled dynamic arrays, cell
metadata, slicers, rich data and VBA alive: they are never parsed, so they
cannot be lost.

The modules are deliberately small and independent of the fill logic:

* :mod:`fillxl._ooxml.package` — the ZIP, its relationships and content types
* :mod:`fillxl._ooxml.xmltext` — text-level XML helpers (no re-serialising)
* :mod:`fillxl._ooxml.workbook` — sheets, defined names, calculation settings
* :mod:`fillxl._ooxml.worksheet` — rows, cells, merged ranges, array ranges
* :mod:`fillxl._ooxml.cells` — a value as a ``<c>`` element
* :mod:`fillxl._ooxml.styles` — number formats, including adding a date format
* :mod:`fillxl._ooxml.tables` — table parts
* :mod:`fillxl._ooxml.charts` — classic chart references
* :mod:`fillxl._ooxml.pivots` — pivot cache refresh flags
* :mod:`fillxl._ooxml.dangling` — references left pointing at a deleted sheet
* :mod:`fillxl._ooxml.fill` — the fill semantics that use all of the above
"""
