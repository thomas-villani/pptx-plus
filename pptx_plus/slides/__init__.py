"""Slide lifecycle — delete, move, and duplicate, at the OPC layer.

The capability this library exists for. python-pptx models a presentation's
content well and its package barely at all, and the gap that matters is the
slide lifecycle: its two oldest open feature requests are "delete a slide" and
"duplicate a slide", both roughly twelve years old.

The gap is not a missing convenience method. Doing any of it correctly means
rewriting relationships and parts, and every recipe in circulation skips
exactly that step (SPEC §3.6).

This package depends only on :mod:`pptx_plus.core`, enforced by
``tests/test_import_invariant.py``. SPEC §5.
"""

from __future__ import annotations

from pptx_plus.slides.delete import delete_slide
from pptx_plus.slides.duplicate import duplicate_slide
from pptx_plus.slides.move import move_slide
from pptx_plus.slides.resolve import (
    SlideIndexError,
    SlideNotFoundError,
    contains,
    resolve_slide,
    slide_index,
)

__all__ = [
    "SlideIndexError",
    "SlideNotFoundError",
    "contains",
    "delete_slide",
    "duplicate_slide",
    "move_slide",
    "resolve_slide",
    "slide_index",
]
