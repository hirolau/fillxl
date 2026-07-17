"""fillxl — Fill Excel templates with data from Python."""

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

__version__ = "0.2.0-beta.1"
