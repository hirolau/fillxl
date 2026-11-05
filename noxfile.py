"""Nox sessions for fillxl multi-version testing.

Tests every major version of each optional data backend from the minimum
supported version to the latest, ensuring backward compatibility.

Version-matrix sessions install the *minimum* supported version, the *latest*
version from ``uv.lock``, and any explicitly pinned majors (currently polars 2
and duckdb 2) separately to catch regressions.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, replace
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

    args = ["pytest", "tests/", "-m", "not win32com"]
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
        "not win32com",
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
# specific major-version range with ``uv pip install`` so the environment (but
# not the lock file) reflects the version we actually want to test.


@dataclass(frozen=True)
class _BackendSession:
    """One pinned-version session for an optional data backend.

    Attributes:
        label: Session name suffix, e.g. ``polars-2`` → ``tests_polars-2``.
        package: Distribution to pin; also the extra that installs it.
        pin: Version constraint to install after sync; empty means "uv.lock".
        select: ``-k`` keyword.  ``-k`` also matches markers, so tests that
            exercise the backend carry the ``polars``/``duckdb`` marker and
            are picked up even when their name never mentions it.
        unavailable_on: Pythons this pin cannot even install on (no wheels).

    """

    label: str
    package: str
    pin: str
    select: str
    unavailable_on: tuple[str, ...] = ()


_DATA_BACKENDS: list[_BackendSession] = [
    _BackendSession("pandas", "pandas", ">=2.0,<3", "pandas"),
    _BackendSession("polars", "polars", ">=0.20,<1", "polars"),
    # duckdb 1.x publishes no cp315 wheels; duckdb 2.x is abi3 and does install.
    _BackendSession("duckdb", "duckdb", ">=1.0,<2", "duckdb", unavailable_on=("3.15",)),
]


def _make_version_session(spec: _BackendSession) -> Any:  # noqa: ANN401
    """Return a nox session function pinned to a specific dependency version."""

    @nox.session(python=PYTHONS, name=f"tests_{spec.label}")
    def version_session(session: nox.Session) -> None:
        if session.python in spec.unavailable_on:
            session.skip(f"{spec.package}{spec.pin} has no wheels for Python {session.python}")
        _run_tests(
            session,
            extras=spec.package,
            pin={spec.package: spec.pin},
            pytest_args=["-k", f"{spec.select} or validation"],
        )

    return version_session


# Generate min + latest sessions for each backend
for _backend in _DATA_BACKENDS:
    _make_version_session(replace(_backend, label=f"{_backend.label}-min", pin=_backend.pin))
    _make_version_session(replace(_backend, label=f"{_backend.label}-latest", pin=""))

# Convenience aliases without suffix → latest
for _backend in _DATA_BACKENDS:
    _make_version_session(replace(_backend, pin=""))

# Additional major-version sessions for backends with more than one supported
# release line.  The "-latest" sessions follow uv.lock; pinning a major here
# keeps that version covered even when the lock moves on.
_EXTRA_MAJOR_SESSIONS: list[_BackendSession] = [
    _BackendSession("polars-2", "polars", ">=2,<3", "polars"),
    _BackendSession("duckdb-2", "duckdb", ">=2,<3", "duckdb"),
]

for _backend in _EXTRA_MAJOR_SESSIONS:
    _make_version_session(_backend)
