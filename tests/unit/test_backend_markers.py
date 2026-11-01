"""The pinned-version sessions select tests by marker; keep that selection honest.

``nox -s tests_polars-2`` and ``nox -s tests_duckdb-2`` select their tests with
``pytest -k "polars or validation"``, and ``-k`` matches marker names as well
as test names.  A test that exercises polars or duckdb therefore has to carry
the matching marker: otherwise it silently drops out of every pinned-version
run and the matrix stays green while covering nothing.

The check below walks the collected tests, so it sees every test only when the
whole suite is collected (``pytest`` or ``nox -s tests``), which is how CI and
the nox matrix entry point run it.
"""

from __future__ import annotations

import pytest

_BACKENDS = ("polars", "duckdb")


@pytest.mark.parametrize("backend", _BACKENDS)
def test_every_test_that_needs_a_backend_is_marked(
    request: pytest.FixtureRequest,
    backend: str,
) -> None:
    """A test calling ``importorskip(backend)`` also carries the backend marker."""
    unmarked: list[str] = []
    seen: set[int] = set()

    for item in request.session.items:
        function = getattr(item, "function", None)
        if function is None or id(function) in seen:
            continue
        seen.add(id(function))

        # ``pytest.importorskip("polars")`` keeps the module name as a constant.
        if backend not in function.__code__.co_consts:
            continue
        if item.get_closest_marker(backend) is None:
            unmarked.append(f"{function.__module__}::{function.__qualname__}")

    assert not unmarked, (
        f"these tests use {backend} without @pytest.mark.{backend}, so the "
        f"pinned-version sessions skip them: {sorted(unmarked)}"
    )
