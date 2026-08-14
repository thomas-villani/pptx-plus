"""Relationship- and content-type constants, and the clone policy sets.

python-pptx's ``RELATIONSHIP_TYPE`` and ``CONTENT_TYPE`` enums cover almost
everything this library needs. The gaps are filled here rather than inline, so
there is one place to look when a new part type turns up in a real deck.

The more important half of this module is the two **policy sets**. They encode
the share-vs-deep rule that drives :mod:`pptx_plus.core.clone`, as data rather
than as branches in the engine — which is what keeps the engine free of
per-part-type special cases. SPEC §4.3.

This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from pptx_plus.core.ns import qn

# ---------------------------------------------------------------------------
# Constants python-pptx omits.
# ---------------------------------------------------------------------------

#: SmartArt's **fifth** part: the rendered drawing cache holding the positioned
#: shapes every viewer actually draws.
#:
#: python-pptx has the four ECMA-376 diagram relationship types
#: (``DIAGRAM_DATA``, ``DIAGRAM_LAYOUT``, ``DIAGRAM_COLORS``,
#: ``DIAGRAM_QUICK_STYLE``) but not this one, because it is a Microsoft
#: extension rather than part of the base standard. Missing it means a cloned
#: diagram loses its cache — PowerPoint recovers by recomputing on open, but
#: every other renderer draws nothing at all.
RT_DIAGRAM_DRAWING = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"

#: ``<p:ext uri=…>`` for the PowerPoint sections extension, holding
#: ``p14:sectionLst``. Sections are keyed on ``sldId/@id``, python-pptx does
#: not model them, and a stale entry produces a repair prompt. SPEC §3.4.
EXT_URI_SECTION_LST = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

#: ``<a:ext uri=…>`` for the SmartArt data-model extension, holding
#: ``dsp:dataModelExt`` and its unqualified ``@relId``.
EXT_URI_DATA_MODEL = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"


#: Scope value: the id resolves against the relationships of the part that
#: contains it, like every ``r:``-namespace attribute.
SCOPE_SELF = "self"

#: Scope value: the id resolves against the relationships of the part that
#: **refers to** the containing part.
SCOPE_PARENT = "parent"

#: Relationship-id-bearing attributes that are **not** in the ``r:`` namespace,
#: mapping ``(element Clark name, attribute name)`` to the scope its value
#: resolves in.
#:
#: The rewriter's core strategy is a sweep of the ``r:`` namespace, which is
#: closed by schema and therefore has zero false positives. This registry is
#: the documented set of exceptions to that closure — and it exists because of
#: exactly one attribute in real decks.
#:
#: Inside a SmartArt data part, ``dgm:extLst/a:ext/dsp:dataModelExt/@relId``
#: points at the ``dsp:`` rendered-drawing part. Miss it and PowerPoint
#: silently recovers by recomputing the drawing on open, while every other
#: renderer draws nothing at all — a failure invisible in the one viewer most
#: likely to be used to check.
#:
#: **Its scope is the referring part, not its own**, and that is the genuinely
#: surprising half. The diagram data part has no relationships at all; its
#: ``relId`` names a relationship on the *slide* that references the diagram.
#: This is verified against a PowerPoint-authored sample rather than assumed —
#: see ``tests/fixtures/pptx_samples/README.md``. Every other relationship
#: reference in OOXML is part-scoped (SPEC §3.2), so code that assumes the
#: universal rule silently produces a copy whose diagram points at the
#: *original's* drawing cache.
#:
#: The registry is deliberately not a longer list of guesses. Anything *not*
#: here that still looks like a relationship id is caught by the unclaimed-
#: literal detector (SPEC §4.4), which is how the next one gets found.
UNQUALIFIED_REL_ID_ATTRS: dict[tuple[str, str], str] = {
    (qn("dsp:dataModelExt"), "relId"): SCOPE_PARENT,
}

#: Relationship types whose **target part's content** holds parent-scoped
#: relationship ids — i.e. ids naming relationships on the part that refers to
#: it, not on itself.
#:
#: The clone engine has to treat these specially: the target cannot be copied
#: until the referring part's relationship map is complete, because rewriting
#: its contents needs that map. Expressed as data so the engine keeps no
#: per-part-type branch. SPEC §4.5.
PARENT_SCOPED_CONTENT_RELTYPES: frozenset[str] = frozenset({RT.DIAGRAM_DATA})


# ---------------------------------------------------------------------------
# The clone policy sets — SPEC §4.5.
# ---------------------------------------------------------------------------

#: Relationship types whose targets may be **shared by reference** rather than
#: cloned: the part's identity is its bytes, and two slides pointing at one
#: part is the format's own way of saying "the same picture."
#:
#: Membership here is **necessary but not sufficient**. The clone engine also
#: requires that the target have zero relationships of its own, because
#: anything owning a sub-graph is a mutable unit whatever this set says. That
#: second clause is the safety net: if a future Office release ships an image
#: part with a sub-relationship, it gets cloned rather than silently aliased.
SHARE_RELTYPES: frozenset[str] = frozenset(
    {
        RT.IMAGE,
        RT.MEDIA,
        RT.VIDEO,
        RT.AUDIO,
        RT.FONT,
        RT.THUMBNAIL,
    }
)

#: Relationship types pointing **up** the deck's structure rather than down
#: into a slide's own content. Never cloned within one deck: two slides using
#: one layout is correct, and duplicating the layout would be wrong.
#:
#: This set is the seam for cross-deck copy (SPEC §6). A same-deck operation
#: reuses these targets verbatim; a cross-deck one either imports the chain
#: into the target or remaps it onto an existing target layout. That choice
#: lives in one resolver function, and nothing else in the engine changes.
STRUCTURAL_RELTYPES: frozenset[str] = frozenset(
    {
        RT.SLIDE_LAYOUT,
        RT.SLIDE_MASTER,
        RT.NOTES_MASTER,
        RT.HANDOUT_MASTER,
        RT.THEME,
    }
)

#: Relationship types naming another *slide* in the same deck. Never cloned:
#: duplicating a slide that links to slide 4 must produce a slide that links to
#: slide 4, not one that drags a second copy of slide 4 along with it.
#:
#: Kept separate from :data:`STRUCTURAL_RELTYPES` deliberately, even though
#: both mean "reuse the target" today. Structural relationships are the
#: cross-deck seam (SPEC §6), where the layout chain gets imported into the
#: target deck; a slide-jump target has entirely different cross-deck
#: semantics, and folding the two together would silently give it the wrong
#: ones the moment §6 lands.
#:
#: This does **not** cover the notes slide's relationship *back* to its own
#: slide, which is a cycle rather than a reference to a third party. The clone
#: map resolves that: by the time the notes slide is cloned its source slide is
#: already mapped, so the back-reference lands on the clone. SPEC §4.5.
REUSE_RELTYPES: frozenset[str] = frozenset({RT.SLIDE})

#: The four ECMA-376 SmartArt definition relationship types plus the Microsoft
#: drawing-cache extension. Not consulted by the clone engine — they are all
#: deep-cloned by the default rule, with no special case — but named here
#: because the test suite asserts all five survive a duplicate, and a list
#: spelled out in a test drifts from one spelled out in code.
DIAGRAM_RELTYPES: frozenset[str] = frozenset(
    {
        RT.DIAGRAM_DATA,
        RT.DIAGRAM_LAYOUT,
        RT.DIAGRAM_COLORS,
        RT.DIAGRAM_QUICK_STYLE,
        RT_DIAGRAM_DRAWING,
    }
)


# ---------------------------------------------------------------------------
# Part-name templates.
# ---------------------------------------------------------------------------

#: Content type -> ``next_partname`` template for cloned parts.
#:
#: A part type absent from this table still clones: the allocator derives a
#: template from the source part's own name, so an unrecognized part lands
#: beside its original rather than failing. The table exists so the common
#: types land in the conventional place a human would expect when they unzip
#: the result.
PARTNAME_TEMPLATES: dict[str, str] = {
    CT.PML_SLIDE: "/ppt/slides/slide%d.xml",
    CT.PML_NOTES_SLIDE: "/ppt/notesSlides/notesSlide%d.xml",
    CT.DML_CHART: "/ppt/charts/chart%d.xml",
    CT.SML_SHEET: "/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx",
    CT.DML_DIAGRAM_DATA: "/ppt/diagrams/data%d.xml",
    CT.DML_DIAGRAM_LAYOUT: "/ppt/diagrams/layout%d.xml",
    CT.DML_DIAGRAM_COLORS: "/ppt/diagrams/colors%d.xml",
    CT.DML_DIAGRAM_STYLE: "/ppt/diagrams/quickStyle%d.xml",
    CT.DML_DIAGRAM_DRAWING: "/ppt/diagrams/drawing%d.xml",
}


__all__ = [
    "DIAGRAM_RELTYPES",
    "EXT_URI_DATA_MODEL",
    "EXT_URI_SECTION_LST",
    "PARENT_SCOPED_CONTENT_RELTYPES",
    "PARTNAME_TEMPLATES",
    "REUSE_RELTYPES",
    "SCOPE_PARENT",
    "SCOPE_SELF",
    "RT_DIAGRAM_DRAWING",
    "SHARE_RELTYPES",
    "STRUCTURAL_RELTYPES",
    "UNQUALIFIED_REL_ID_ATTRS",
]
