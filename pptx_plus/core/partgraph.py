"""Reading and classifying a part's outgoing relationships.

Read-only. Nothing here mutates a package; the clone engine consumes what this
module reports and does the mutating.

Two jobs. :func:`rel_edges` snapshots a part's relationships into frozen
:class:`RelEdge` records, so the rest of the library never handles a
``_Relationship`` object directly and an upstream change to that class has one
call site to fix. :func:`classify` decides, per edge, whether the clone should
point at the same part or get its own copy.

SPEC §4.5. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import TYPE_CHECKING

from pptx_plus.core.reltypes import REUSE_RELTYPES, SHARE_RELTYPES, STRUCTURAL_RELTYPES

if TYPE_CHECKING:
    from pptx.opc.package import Part


class Disposition(Enum):
    """What the clone should do with one relationship."""

    #: Point at the same part. Its identity is its bytes, and two slides
    #: referencing one image is the format's own way of saying "the same
    #: picture".
    SHARE = "share"
    #: Point at the same part, because it is part of the deck's structure
    #: rather than of this slide's content. The v0.2 cross-deck seam.
    STRUCTURAL = "structural"
    #: Point at the same part, because it names another slide. Duplicating a
    #: slide that links to slide 4 must produce a slide that links to slide 4,
    #: not one that brings a second copy of slide 4 with it.
    REUSE = "reuse"
    #: Give the clone its own copy, recursively.
    DEEP = "deep"
    #: Re-mint by reference. There is no part to clone.
    EXTERNAL = "external"


@dataclass(frozen=True)
class RelEdge:
    """One outgoing relationship, snapshotted away from python-pptx's object.

    Frozen and plain so that ``_Relationship``'s surface leaks no further than
    :func:`rel_edges`. Everything downstream reads these fields.
    """

    r_id: str
    reltype: str
    is_external: bool
    #: The target part, or None when external.
    target_part: Part | None
    #: The raw target string. Only meaningful when external.
    target_ref: str


def rel_edges(part: Part) -> list[RelEdge]:
    """Snapshot a part's outgoing relationships.

    Args:
        part: The part to read.

    Returns:
        One :class:`RelEdge` per relationship, in the collection's own order.

    Returned as a list, not a generator: the caller mints relationships on the
    clone while iterating, and on a same-package clone that is a mutation of a
    collection sharing a package with the one being read.

    Note:
        This iterates ``part.rels`` and touches nothing else. It must stay
        that way. ``SlidePart.notes_slide`` is a lazy property that *creates*
        a notes slide, reaching ``PresentationPart.notes_master_part``, which
        creates a notes master **and** a theme part — so an inspection path
        that touched it would silently grow the package it was only supposed
        to read. SPEC §9.8.
    """
    return [
        RelEdge(
            r_id=rel.rId,
            reltype=rel.reltype,
            is_external=rel.is_external,
            target_part=None if rel.is_external else rel.target_part,
            target_ref=rel.target_ref if rel.is_external else "",
        )
        for rel in part.rels.values()
    ]


def classify(edge: RelEdge) -> Disposition:
    """Decide what a clone should do with one relationship.

    Args:
        edge: The relationship to classify.

    Returns:
        The :class:`Disposition` for this edge.

    The rule, stated once rather than as a table of part types:

        A related part is **shared by reference** iff its relationship type is
        in :data:`~pptx_plus.core.reltypes.SHARE_RELTYPES` — its identity is
        its bytes — **and** it has no relationships of its own.

    The second clause is the safety net, and it does the real work. Anything
    owning a sub-graph is a mutable unit and cannot be shared whatever the
    reltype table says, so if a future Office release ships an image part with
    a sub-relationship it gets cloned rather than silently aliased. Errs toward
    copying, which costs bytes; the other direction costs correctness.

    **Charts need no special case.** A ``ChartPart`` is an ordinary XML part
    whose ``c:chartSpace/c:externalData/@r:id`` points at an
    ``EmbeddedXlsxPart`` — a blob part with no relationships whose reltype is
    ``RT.PACKAGE``, which is not in ``SHARE_RELTYPES``. So it deep-clones
    byte-identically under a fresh name, its relationship is re-minted, and
    its ``@r:id`` is rewritten because it lives in the ``r:`` namespace. The
    "fiddly bit" dissolves into the general rule.

    Example:
        >>> from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        >>> classify(RelEdge("rId1", RT.SLIDE_LAYOUT, False, None, "")).name
        'STRUCTURAL'
    """
    if edge.is_external:
        return Disposition.EXTERNAL
    if edge.reltype in STRUCTURAL_RELTYPES:
        return Disposition.STRUCTURAL
    if edge.reltype in REUSE_RELTYPES:
        return Disposition.REUSE
    if edge.reltype in SHARE_RELTYPES and not _has_relationships(edge.target_part):
        return Disposition.SHARE
    return Disposition.DEEP


def _has_relationships(part: Part | None) -> bool:
    """Return whether a part owns any relationships of its own."""
    return part is not None and len(part.rels) > 0


#: Dispositions that reuse the source's target rather than cloning it.
REUSING = frozenset({Disposition.SHARE, Disposition.STRUCTURAL, Disposition.REUSE})


__all__ = [
    "REUSING",
    "Disposition",
    "RelEdge",
    "classify",
    "rel_edges",
]
