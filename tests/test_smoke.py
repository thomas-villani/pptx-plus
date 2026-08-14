"""Scaffolding smoke tests — the Phase 0 gate.

These exist so CI has something real to grade before any capability code is
written. `test_python_pptx_opens_its_default_template` is the one that earns
its place: on every matrix leg it proves the dependency resolved, that its
bundled `default.pptx` unzips, and — on the min-deps leg — that the declared
lower bound actually works on the oldest supported interpreter.
"""

from __future__ import annotations

import pptx_plus


def test_package_exposes_a_version() -> None:
    assert pptx_plus.__version__


def test_base_error_is_exported() -> None:
    assert issubclass(pptx_plus.PptxPlusError, Exception)


def test_python_pptx_opens_its_default_template() -> None:
    from pptx import Presentation

    assert len(Presentation().slide_layouts) > 0
