"""Chapter-based narrative tests for fillxl.

Each module in this package is a "chapter" that builds on the previous one,
documenting fillxl features through executable tests.  The numeric prefix
ensures pytest collects them in order.

Chapters
--------
01 — Basic fill() behaviour
02 — Filling named cells
03 — Filling Excel tables
04 — Data formats for tables
05 — Configuring fillxl
06 — Column validation
07 — Charts and table references
08 — Sheet management
09 — Advanced features (calculated columns, CF, pivots)
10 — Under the hood: backends
11 — Data types
12 — Template preservation
13 — Failure recovery
14 — Real report generation
15 — Custom Excel formats (.xlsm, .xlsb)
"""

import pytest

pytestmark = pytest.mark.story
