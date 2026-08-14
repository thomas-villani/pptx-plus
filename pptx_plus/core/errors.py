"""Library base error class, isolated to break an import cycle.

:class:`PptxPlusError` lives here — in a module that imports nothing else
from ``pptx_plus.core`` — so the other ``core`` submodules (``ns``,
``ids``, ``relmap``, …) can subclass it with a plain top-of-file import
instead of the ``# noqa: E402`` ordering dance that an in-``__init__``
definition would force. ``core/__init__`` re-exports it, so the documented
short form ``from pptx_plus.core import PptxPlusError`` is unchanged.
SPEC §9.7.
"""

from __future__ import annotations


class PptxPlusError(Exception):
    """Base class for all library-raised errors.

    Every typed error in pptx_plus subclasses this so callers can catch the
    library's failures without catching unrelated ``ValueError`` /
    ``KeyError`` instances raised by python-pptx or lxml.

    Most subclasses additionally inherit a stdlib exception type where it
    aids ``except`` ergonomics — :class:`~pptx_plus.slides.resolve.SlideNotFoundError`
    is also a ``KeyError``, for instance — so existing caller code keeps
    working while ``except PptxPlusError`` catches everything this library
    raises. SPEC §16.
    """


__all__ = ["PptxPlusError"]
