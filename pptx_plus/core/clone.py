"""The part-graph clone engine.

Given a root part, produce a copy of it and of everything it owns, with every
relationship re-minted and every relationship id inside the copied XML
rewritten to match. This is what the naive recipes skip, and it is the whole
of the difference between a duplicate that works and one whose pictures
silently vanish.

The engine has no per-part-type branches. Share-vs-deep is data
(:mod:`pptx_plus.core.reltypes`), classification is one function
(:mod:`pptx_plus.core.partgraph`), and cloning is one expression
(:mod:`pptx_plus.core.parts`). Charts, SmartArt, embedded workbooks and media
all fall out of the general rule — see :func:`~pptx_plus.core.partgraph.classify`.

**Three invariants hold this together.**

*The clone-map invariant.* Within one operation each source part maps to at
most one destination, and any edge whose target is already mapped resolves to
that mapping. This is what makes the notes slide correct for free: a
``NotesSlidePart`` holds a relationship *back* to its slide, and by the time it
is cloned the source slide is already in the map, so the back-reference lands
on the clone rather than on the original. The same mechanism handles diamonds
(two graphic frames on one chart) and any other cycle.

*Map before recurse.* ``part_map[id(src)] = dst`` is assigned **before** the
edges are walked. Assigning afterwards would hang on any deck with speaker
notes, because slide -> notesSlide -> slide is a real cycle.

*Identity, not equality.* The map is keyed on ``id(part)``. ``Part`` defines no
``__eq__``, so equality would be identity anyway — but saying ``id()`` means it
cannot quietly change if that stops being true.

SPEC §4.5. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, TypeGuard

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart

from pptx_plus.core.oxml import part_root
from pptx_plus.core.partgraph import REUSING, Disposition, RelEdge, classify, rel_edges
from pptx_plus.core.parts import clone_part
from pptx_plus.core.relmap import RelMap, remap_rel_ids
from pptx_plus.core.reltypes import PARENT_SCOPED_CONTENT_RELTYPES, UNQUALIFIED_REL_ID_ATTRS

if TYPE_CHECKING:
    from pptx.opc.package import Part
    from pptx.package import Package

#: Cheap pre-filter: only reparse a blob when it plausibly holds one of the
#: registered unqualified attributes. Keeps the exception in SPEC §8.1's
#: byte-exactness as narrow as it can be -- a diagram data part that carries no
#: such attribute is still copied byte for byte.
_CONTAINS_UNQUALIFIED_ID = re.compile(
    b"|".join(re.escape(attr.encode()) for _, attr in sorted(UNQUALIFIED_REL_ID_ATTRS))
)


@dataclass(frozen=True)
class ClonePolicy:
    """Options for one clone operation.

    Attributes:
        with_notes: Clone the notes slide along with its slide.
        strict: Raise :class:`~pptx_plus.core.relmap.DanglingRelationshipError`
            on a relationship id in the source XML that names no relationship.

    ``with_notes`` defaults to True because PowerPoint's own Duplicate Slide
    keeps speaker notes, and because the costs are asymmetric: losing notes is
    unrecoverable, while gaining notes you did not want is one ``del`` away.
    """

    with_notes: bool = True
    strict: bool = True


@dataclass
class CloneResult:
    """What one clone operation produced.

    Attributes:
        root: The clone of the part the operation started from.
        part_map: ``id(source part) -> clone``, for every part cloned.
        shared: Parts the clone points at rather than copying.
    """

    root: Part
    part_map: dict[int, Part] = field(default_factory=dict)
    shared: list[Part] = field(default_factory=list)

    @property
    def cloned_parts(self) -> list[Part]:
        """Every part this operation created, root first."""
        return [self.root, *(p for p in self.part_map.values() if p is not self.root)]


def clone_part_graph(
    root: Part,
    *,
    into: Package,
    policy: ClonePolicy | None = None,
) -> CloneResult:
    """Clone a part and everything it owns into a package.

    Args:
        root: The part to clone — a slide, for every v0.1 caller.
        into: The destination package. Always the source's own package at
            v0.1; the parameter exists so cross-deck copy (SPEC §6) changes
            two functions rather than this signature.
        policy: Clone options. Defaults to :class:`ClonePolicy`.

    Returns:
        A :class:`CloneResult`. The clone is **not attached to anything** —
        minting the relationship that puts it in the deck is the caller's job,
        because only the caller knows what it should hang off.

    Raises:
        DanglingRelationshipError: Under ``policy.strict``, the source XML
            names a relationship that does not exist.

    The clone has real relationships of its own, minted in the source's
    iteration order — but **not necessarily with the same ids**.
    ``_Relationships._next_rId`` fills gaps in the numbering, and
    ``get_or_add`` dedupes on ``(reltype, target, is_external)``, so a source
    whose relationships are ``{rId1, rId3}`` — the ordinary result of deleting
    a shape in PowerPoint — produces a clone with ``{rId1, rId2}``. That is
    exactly why the rewrite is not optional, and exactly why a test built on a
    freshly authored deck will not notice if you skip it: in the simple case
    the map comes out as the identity.
    """
    policy = policy or ClonePolicy()
    reserved: set[str] = set()
    part_map: dict[int, Part] = {}
    shared: list[Part] = []

    def _clone(src: Part) -> Part:
        existing = part_map.get(id(src))
        if existing is not None:
            return existing

        dst = clone_part(src, into=into, reserved=reserved)
        # Before walking the edges, not after: slide -> notesSlide -> slide is
        # a real cycle and this assignment is what terminates it.
        part_map[id(src)] = dst

        relmap: RelMap = {}
        deferred: list[RelEdge] = []
        for edge in rel_edges(src):
            if not policy.with_notes and edge.reltype == RT.NOTES_SLIDE:
                continue

            # A SmartArt data part's contents name relationships on *this*
            # part, not on itself, so it cannot be copied until `relmap` is
            # finished. Held back rather than special-cased downstream, which
            # keeps the rest of the loop free of the exception. SPEC §4.5.
            if edge.reltype in PARENT_SCOPED_CONTENT_RELTYPES and not edge.is_external:
                deferred.append(edge)
                continue

            disposition = classify(edge)
            if disposition is Disposition.EXTERNAL:
                relmap[edge.r_id] = dst.rels.get_or_add_ext_rel(edge.reltype, edge.target_ref)
                continue

            target = edge.target_part
            assert target is not None  # not external, so a part exists

            # The clone-map invariant, and it must be tested *before* the
            # disposition: a REUSE edge to a slide that is itself being cloned
            # (the notes slide's back-reference) has to land on the clone, not
            # on the source.
            mapped = part_map.get(id(target))
            if mapped is not None:
                relmap[edge.r_id] = dst.relate_to(mapped, edge.reltype)
                continue

            if disposition in REUSING:
                if disposition is Disposition.SHARE:
                    shared.append(target)
                relmap[edge.r_id] = dst.relate_to(target, edge.reltype)
                continue

            relmap[edge.r_id] = dst.relate_to(_clone(target), edge.reltype)

        # `relmap` is complete now, so the held-back parts can be rewritten
        # against it. Nothing inside such a part references the part itself,
        # so minting its own relationship afterwards is safe.
        for edge in deferred:
            target = edge.target_part
            assert target is not None
            child = _clone_parent_scoped(target, relmap)
            relmap[edge.r_id] = dst.relate_to(child, edge.reltype)

        # After every relationship is minted, so the map is complete -- and
        # per cloned XML part, because a chart's `c:externalData/@r:id` lives
        # in a sub-part with a map of its own.
        if _is_xml(dst):
            remap_rel_ids(part_root(dst), relmap, strict=policy.strict)

        return dst

    def _clone_parent_scoped(src: Part, parent_relmap: RelMap) -> Part:
        """Copy a part whose contents name relationships on the referring part.

        python-pptx has no model class for a SmartArt data part, so it loads
        as an opaque blob with no element tree -- which is why the ordinary
        rewrite pass skips it entirely, and why a naive implementation leaves
        the copy's diagram pointing at the *original's* drawing cache with no
        error anywhere. The bytes have to be rewritten before the part is
        loaded, because a loaded blob part has no public way to change them.

        This is the one place the library reparses a blob part, so it is the
        one exception to the byte-exactness of SPEC §8.1 -- and an unavoidable
        one: the reference has to change, so the bytes cannot stay identical.
        """
        blob = src.blob
        if _CONTAINS_UNQUALIFIED_ID.search(blob):
            element = etree.fromstring(blob)
            remap_rel_ids(element, {}, parent=parent_relmap, strict=policy.strict)
            blob = etree.tostring(element, xml_declaration=True, encoding="UTF-8", standalone=True)
        return clone_part(src, into=into, reserved=reserved, blob=blob)

    return CloneResult(root=_clone(root), part_map=part_map, shared=shared)


def _is_xml(part: Part) -> TypeGuard[XmlPart]:
    """Return whether a part has a parsed element tree to rewrite.

    Tested by class rather than by probing for an attribute: ``XmlPart`` is
    public, and reaching for ``_element`` here would put a third private
    accessor outside the quarantine in ``core/oxml.py`` (SPEC §14.2).

    A blob part -- an image, a media file, a SmartArt definition part, an
    embedded workbook -- has no element tree and must not be reparsed. Copying
    it as bytes is what makes the preservation guarantee byte-exact rather
    than approximate (SPEC §8.1).
    """
    return isinstance(part, XmlPart)


__all__ = ["ClonePolicy", "CloneResult", "clone_part_graph"]
