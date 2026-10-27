"""Warning categories fillxl emits, and the helper that points them at user code.

Every warning fillxl raises about the data or the template is a
:class:`FillXLWarning`, so a single filter can turn all of them into errors::

    warnings.simplefilter("error", fillxl.FillXLWarning)

Diagnostics that are not the caller's to act on stay in the log.
"""

from __future__ import annotations

import sys
import warnings
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from types import FrameType


class FillXLWarning(UserWarning):
    """Base class for the warnings fillxl emits."""


class PrecisionLossWarning(FillXLWarning):
    """A value is stored with less precision than it was supplied with."""


class IgnoredDataWarning(FillXLWarning):
    """Part of the input was not written to the output."""


class TimezoneWarning(FillXLWarning):
    """A timezone was stripped: Excel stores no timezone information."""


def warn_user(message: str, category: type[FillXLWarning]) -> None:
    """Warn about *message*, blaming the first frame outside fillxl.

    The warning is raised deep inside the backends, but it belongs to whoever
    called :func:`fillxl.fill`: ``pytest.warns``, ``warnings.filterwarnings``
    and ``-W error`` all match on that frame's module.

    Examples:
        >>> import warnings
        >>> with warnings.catch_warnings(record=True) as caught:
        ...     warnings.simplefilter("always")
        ...     warn_user("example", PrecisionLossWarning)
        >>> caught[0].category.__name__
        'PrecisionLossWarning'
        >>> str(caught[0].message)
        'example'

    """
    frame: FrameType | None = sys._getframe(1)
    stacklevel = 2
    while frame is not None and _is_internal(frame):
        frame = frame.f_back
        stacklevel += 1
    warnings.warn(message, category, stacklevel=stacklevel)


def _is_internal(frame: FrameType) -> bool:
    """Whether *frame* runs inside the fillxl package."""
    module = frame.f_globals.get("__name__")
    return isinstance(module, str) and (module == "fillxl" or module.startswith("fillxl."))
