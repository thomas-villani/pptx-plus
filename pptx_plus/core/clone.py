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

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, TypeGuard

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart

from pptx_plus.core.oxml import part_root
from pptx_plus.core.partgraph import REUSING, Disposition, classify, rel_edges
from pptx_plus.core.parts import clone_part
from pptx_plus.core.relmap import RelMap, remap_rel_ids

if TYPE_CHECKING:
    from pptx.opc.package import Part
    from pptx.package import Package


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
        for edge in rel_edges(src):
            if not policy.with_notes and edge.reltype == RT.NOTES_SLIDE:
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

        # After every relationship is minted, so the map is complete -- and
        # per cloned XML part, because a chart's `c:externalData/@r:id` and a
        # diagram's `@relId` live in sub-parts with maps of their own.
        if _is_xml(dst):
            remap_rel_ids(part_root(dst), relmap, strict=policy.strict)

        return dst

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
