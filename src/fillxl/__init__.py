"""fillxl — Fill Excel templates with data from Python."""

from importlib.metadata import PackageNotFoundError, version

from fillxl._config import FillXLConfig
from fillxl._core import fill
from fillxl._types import Engine
from fillxl._validation import ColumnValidationStrategy

__all__ = [
    "ColumnValidationStrategy",
    "Engine",
    "FillXLConfig",
    "__version__",
    "fill",
]

try:
    __version__ = version("fillxl")
except PackageNotFoundError:  # pragma: no cover - only when running from a source tree
    __version__ = "0.0.0"
