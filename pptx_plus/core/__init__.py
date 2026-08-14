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
from pptx_plus.core.ids import (
    MAX_SLIDE_ID,
    MIN_SLIDE_ID,
    SlideIdRangeError,
    next_slide_id,
    used_slide_ids,
    validate_slide_id,
)
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
from pptx_plus.core.parts import (
    UnclonablePartError,
    allocate_partname,
    clone_part,
    drop_relationship,
    partname_template_for,
)
from pptx_plus.core.reltypes import (
    DIAGRAM_RELTYPES,
    EXT_URI_DATA_MODEL,
    EXT_URI_SECTION_LST,
    PARTNAME_TEMPLATES,
    RT_DIAGRAM_DRAWING,
    SHARE_RELTYPES,
    STRUCTURAL_RELTYPES,
)
from pptx_plus.core.sections import (
    custom_show_lst,
    reorder_slide,
    scrub_slide,
    section_lst,
)

__all__ = [
    "BUILD_NSMAP",
    "DIAGRAM_RELTYPES",
    "EXT_URI_DATA_MODEL",
    "EXT_URI_SECTION_LST",
    "MAX_SLIDE_ID",
    "MIN_SLIDE_ID",
    "NSMAP",
    "PARTNAME_TEMPLATES",
    "RT_DIAGRAM_DRAWING",
    "SHARE_RELTYPES",
    "STRUCTURAL_RELTYPES",
    "InvalidNamespaceError",
    "PptxPlusError",
    "SlideIdRangeError",
    "UnclonablePartError",
    "UpstreamSurfaceError",
    "allocate_partname",
    "check_upstream_surface",
    "clone_part",
    "custom_show_lst",
    "drop_relationship",
    "el",
    "next_slide_id",
    "ordered_insert",
    "part_root",
    "partname_template_for",
    "qn",
    "remove",
    "reorder_slide",
    "scrub_slide",
    "section_lst",
    "sld_id_lst",
    "sub",
    "used_slide_ids",
    "validate_slide_id",
    "xpath",
]
