"""The warning categories fillxl raises, and where they point."""

from __future__ import annotations

import warnings

import pytest

import fillxl
from fillxl import FillXLWarning, LossyTemplateWarning, TimezoneWarning
from fillxl._warnings import warn_user


def test_every_warning_class_is_a_fillxl_warning() -> None:
    """One filter can catch them all: ``simplefilter("error", FillXLWarning)``."""
    for category in (TimezoneWarning, LossyTemplateWarning):
        assert issubclass(category, FillXLWarning)
        assert issubclass(category, UserWarning)


def test_the_warning_classes_are_exported() -> None:
    for name in ("FillXLWarning", "LossyTemplateWarning", "TimezoneWarning"):
        assert name in fillxl.__all__
        assert getattr(fillxl, name) is not None


def test_a_warning_can_be_turned_into_an_error_with_one_filter() -> None:
    with warnings.catch_warnings():
        warnings.simplefilter("error", FillXLWarning)
        with pytest.raises(FillXLWarning):
            warn_user("example")
