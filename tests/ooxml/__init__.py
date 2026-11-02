"""Tests for the ``engine="fillxl"`` OOXML editors.

The helpers here build small templates with openpyxl and patch their parts as
text, because the shapes that break a surgical editor (shared formulas, cells
without a reference, the 1904 date system, spills) cannot be produced by
openpyxl itself.
"""
