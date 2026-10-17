"""Nox sessions for fillxl multi-version testing.

Tests every major version of each optional data backend from the minimum
supported version to the latest, ensuring backward compatibility.

Version-matrix sessions install the *minimum* supported major version and the
*latest* available version separately to catch regressions.
"""

from __future__ import annotations

import os
from typing import Any

import nox

nox.options.default_venv_backend = "uv"
nox.options.reuse_existing_virtualenvs = True

# Avoid hardlink failures when cache and .nox are on different drives
os.environ.setdefault("UV_LINK_MODE", "copy")

PYTHONS = ["3.11", "3.12", "3.13", "3.14", "3.15"]
_ENV = {"UV_LINK_MODE": "copy"}


# ── Helpers ──────────────────────────────────────────────────────────────────


def _run_tests(
    session: nox.Session,
    *,
    extras: str | list[str] | None = None,
    pin: dict[str, str] | None = None,
    pytest_args: list[str] | None = None,
) -> None:
    """Sync dependencies, optionally pin specific versions, and run pytest.

    Args:
        session: Nox session.
        extras: One or more ``--extra`` values for ``uv sync``.
        pin: Mapping of ``package -> version_constraint`` to install after sync
            (e.g. ``{"pandas": ">=2.0,<3"}``).
        pytest_args: Additional arguments forwarded to pytest.
    """
    cmd = ["uv", "sync", "--only-dev"]
    if extras:
        if isinstance(extras, str):
            cmd += ["--extra", extras]
        else:
            for extra in extras:
                cmd += ["--extra", extra]
    session.run_always(*cmd, external=True, env=_ENV)

    # Pin specific versions if requested (overrides the loose extra constraint)
    # Uses ``uv pip install`` so pyproject.toml / uv.lock are NOT modified.
    if pin:
        for pkg, constraint in pin.items():
            if constraint:
                session.run("uv", "pip", "install", f"{pkg}{constraint}", env=_ENV)

    args = ["pytest", "tests/", "-m", "not xlwings"]
    if pytest_args:
        args += pytest_args
    session.run(*args, env=_ENV)


# ── Core sessions ────────────────────────────────────────────────────────────


@nox.session(python=PYTHONS)
def tests(session: nox.Session) -> None:
    """Full test suite with all extras at latest versions + coverage."""
    session.run_always("uv", "sync", "--all-extras", external=True, env=_ENV)
    session.run(
        "pytest",
        "tests/",
        "-m",
        "not xlwings",
        "--cov=fillxl",
        "--cov-report=term-missing",
        "--cov-report=xml",
        env={**_ENV, "COVERAGE_FILE": f".coverage.{session.python}"},
    )


@nox.session(python=PYTHONS)
def tests_minimal(session: nox.Session) -> None:
    """No optional data backends — only list[dict] / dict[str,list] paths."""
    _run_tests(session, pytest_args=["-k", "validation"])


@nox.session(python=PYTHONS)
def mypy(session: nox.Session) -> None:
    """Strict type checking with mypy."""
    session.run_always("uv", "sync", "--all-extras", external=True, env=_ENV)
    session.run("mypy", "src/fillxl/", env=_ENV)


@nox.session(python=PYTHONS)
def lint(session: nox.Session) -> None:
    """Ruff linter + formatter checks."""
    session.run_always("uv", "sync", "--all-extras", external=True, env=_ENV)
    session.run("ruff", "check", "src/fillxl/", "tests/", env=_ENV)
    session.run("ruff", "format", "--check", "src/fillxl/", "tests/", env=_ENV)


@nox.session(python=PYTHONS)
def pyrefly(session: nox.Session) -> None:
    """Type checking with Meta's Pyrefly."""
    session.run_always("uv", "sync", "--all-extras", external=True, env=_ENV)
    session.run("pyrefly", "check", env=_ENV)


# ── Version-matrix sessions ─────────────────────────────────────────────────
#
# For each optional data library we test the MINIMUM supported major version
# and the LATEST major version.  This catches regressions that slip in when
# only the latest version is exercised.
#
# We install the base extra first (e.g. "pandas" → pandas>=2.0), then pin a
# specific major-version range with ``uv add`` so the lock file reflects the
# version we actually want to test.

_DATA_BACKENDS: list[tuple[str, str, str, str]] = [
    # (label, extra, min_pin, pytest_keyword)
    ("pandas", "pandas", ">=2.0,<3", "pandas"),
    ("polars", "polars", ">=0.20,<1", "polars"),
    ("duckdb", "duckdb", ">=1.0,<2", "duckdb"),
]


def _make_version_session(label: str, extra: str, pin: str, keyword: str) -> Any:  # noqa: ANN401
    """Return a nox session function pinned to a specific dependency version."""

    @nox.session(python=PYTHONS, name=f"tests_{label}")
    def version_session(session: nox.Session) -> None:
        _run_tests(
            session,
            extras=extra,
            pin={keyword: pin},
            pytest_args=["-k", f"{keyword} or validation"],
        )

    return version_session


# Generate min + latest sessions for each backend
for _label, _extra, _min_pin, _keyword in _DATA_BACKENDS:
    _make_version_session(f"{_label}-min", _extra, _min_pin, _keyword)
    _make_version_session(f"{_label}-latest", _extra, "", _keyword)

# Convenience aliases without suffix → latest
for _label, _extra, _min_pin, _keyword in _DATA_BACKENDS:
    _make_version_session(_label, _extra, "", _keyword)
