"""Foundation layer — the only module capability packages may import from.

``core`` owns everything that touches the OPC package directly: namespaces
(:mod:`~pptx_plus.core.ns`), element construction
(:mod:`~pptx_plus.core.oxml`), relationship and part manipulation, and the
part-graph clone engine. It depends on nothing above it, and
``tests/test_import_invariant.py`` enforces that capability packages depend
on nothing *but* it. SPEC §9.1.

Symbols are re-exported here so callers can use the short form
``from pptx_plus.core import PptxPlusError`` as well as the long form
``from pptx_plus.core.errors import PptxPlusError``. Both are supported; the
short form is what the docs use.

.. note::
   **There is no orphan collector here, deliberately.**
   ``OpcPackage.save`` serializes ``tuple(self.iter_parts())``, and
   ``iter_parts`` is a relationship-graph walk from the package root — so a
   part that nothing references is simply never written, and
   ``[Content_Types].xml`` is regenerated from the same tuple. Dropping a
   slide's relationship is the whole of "collecting" it. A hand-rolled
   mark-and-sweep would be a second, weaker implementation of something the
   writer already does correctly. SPEC §3.5.
"""

from __future__ import annotations

from pptx_plus.core._compat import UpstreamSurfaceError, check_upstream_surface
from pptx_plus.core.errors import PptxPlusError
from pptx_plus.core.ns import (
    BUILD_NSMAP,
    NSMAP,
    InvalidNamespaceError,
    qn,
)
from pptx_plus.core.oxml import (
    el,
    ordered_insert,
    part_root,
    remove,
    sld_id_lst,
    sub,
    xpath,
)

__all__ = [
    "BUILD_NSMAP",
    "NSMAP",
    "InvalidNamespaceError",
    "PptxPlusError",
    "UpstreamSurfaceError",
    "check_upstream_surface",
    "el",
    "ordered_insert",
    "part_root",
    "qn",
    "remove",
    "sld_id_lst",
    "sub",
    "xpath",
]
