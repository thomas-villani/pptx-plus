"""Runnable examples, shipped in the wheel.

Each module runs standalone::

    python -m pptx_plus.examples.slide_lifecycle

Omitted from coverage and exempt from docstring linting. Output is **ASCII
only**: these run on a default Windows console, whose code page is cp1252, and
a stray typographic quote or arrow raises `UnicodeEncodeError` on the print
rather than anywhere near the code that produced it. SPEC §11.
"""

from __future__ import annotations

__all__: list[str] = []
