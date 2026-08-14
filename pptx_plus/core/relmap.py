"""Rewriting relationship ids inside cloned XML.

The single most important module in this library, because it is the step every
circulating "duplicate a slide" recipe omits.

A relationship id is **scoped to the part that contains it** (SPEC §3.2).
``rId2`` in ``slide1.xml`` and ``rId2`` in ``slide2.xml`` are unrelated names
in unrelated namespaces. So copying a slide's XML into a new part carries
every ``r:id`` and ``r:embed`` in it across unchanged, where they now resolve
against the *new* part's relationships — meaning something else, or nothing.
The picture does not error. It silently fails to render.

Rewriting them needs a map from the source part's ids to the clone's, and a
pass that applies it to every relationship-id-bearing attribute.

**The map is a function, not a bijection.** ``_Relationships.get_or_add``
dedupes on ``(reltype, target, is_external)``, so two source relationships
with the same type and target collapse to one on the clone and two keys map to
one value. Code that assumes bijectivity — inverting the map, asserting equal
lengths — is wrong.

ECMA-376 Part 1 §15.3 (relationship references), and
``shared-relationshipReference.xsd`` for the closed attribute list.

SPEC §4.4. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import re
import warnings
from typing import TYPE_CHECKING

from pptx_plus.core.errors import PptxPlusError
from pptx_plus.core.ns import R
from pptx_plus.core.reltypes import UNQUALIFIED_REL_ID_ATTRS

if TYPE_CHECKING:
    from lxml import etree

#: Source relationship id -> clone relationship id. Many-to-one; see above.
RelMap = dict[str, str]

#: Clark-notation prefix of the ``r:`` namespace.
_R_PREFIX = f"{{{R}}}"

#: What a relationship id looks like, for the unclaimed-literal tripwire.
_RID_LITERAL = re.compile(r"^rId\d+$")


class DanglingRelationshipError(PptxPlusError, ValueError):
    """Raised when a relationship-id attribute names an id not in the map.

    Subclasses ``ValueError`` per SPEC §9.7. Real decks do contain dangling
    references — a shape deleted in a way that left its ``r:embed`` behind, a
    third-party writer that dropped a relationship — and propagating one
    silently into a clone is worse than refusing, because the clone is
    something this library produced and vouches for.
    """


class RelIdLiteralWarning(UserWarning):
    """An attribute holds something shaped like a relationship id, unrewritten.

    The tripwire for a relationship-id attribute this library does not know
    about. It is a ``UserWarning`` rather than an error at runtime because the
    match is heuristic: a shape genuinely named ``rId7`` must not crash a
    library call. In the test suite it is promoted to an error, which is how
    the next ``dsp:dataModelExt/@relId`` gets found rather than silently
    producing a dangling reference in someone's deck.

    A ``UserWarning`` and not a subclass of ``PptxPlusError``: warnings are
    filtered by class, and making it catchable as an error would invite
    exactly the ``except PptxPlusError: pass`` that defeats it.
    """


def _rel_id_attrs(root: etree._Element) -> list[tuple[etree._Element, str, str]]:
    """Collect every ``(element, attribute, value)`` holding a relationship id.

    Two sources, and the asymmetry between them is the design.

    The ``r:`` namespace is swept **wholesale**, because it is closed by
    schema: ``shared-relationshipReference.xsd`` declares exactly nine
    attributes — ``r:id``, ``r:embed``, ``r:link``, ``r:dm``, ``r:lo``,
    ``r:qs``, ``r:cs``, ``r:href``, ``r:pict`` — and every relationship
    reference attribute group in OOXML references one of them. A namespace
    sweep therefore has **zero false positives by construction**, and it picks
    up ``mc:Choice`` and ``mc:Fallback`` content for free.

    The alternative — an allowlist keyed on element name — is keyed on an
    *open* vocabulary. ``a:blip``, ``a:hlinkClick``, ``dgm:relIds``,
    ``p:oleObj``, ``p:videoFile``, ``c:externalData``, ``p:tags``,
    ``v:imagedata``, plus everything under ``mc:AlternateContent``:
    ``a14:imgLayer``, ``p14:media``, ``asvg:svgBlip``. That list grows with
    every Office release and fails **silently** on each addition. It would
    also re-import python-pptx's object model through the back door, which is
    what the fidelity thesis says this library must not do.

    Everything outside the ``r:`` namespace has to be registered by name, in
    :data:`~pptx_plus.core.reltypes.UNQUALIFIED_REL_ID_ATTRS`.

    Returns the results as a list rather than a generator: the caller mutates
    the attributes it is handed, and collecting first keeps the traversal
    separate from the mutation.
    """
    found: list[tuple[etree._Element, str, str]] = []
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue  # a comment or processing instruction
        for key, value in node.attrib.items():
            name = str(key)
            if name.startswith(_R_PREFIX) or (node.tag, name) in UNQUALIFIED_REL_ID_ATTRS:
                found.append((node, name, str(value)))
    return found


def remap_rel_ids(root: etree._Element, relmap: RelMap, *, strict: bool = True) -> int:
    """Rewrite every relationship-id attribute beneath ``root``, in place.

    Args:
        root: The root element of a cloned XML part.
        relmap: Source relationship id -> clone relationship id.
        strict: Raise on an id absent from the map. When False, an unmapped id
            is left alone.

    Returns:
        The number of attributes rewritten.

    Raises:
        DanglingRelationshipError: ``strict`` and an ``r:``-namespace
            attribute holds an id that is not in ``relmap``.

    Warns:
        RelIdLiteralWarning: An attribute outside the known set holds a value
            shaped like a relationship id.

    An **empty value is not a dangling reference.**
    ``<a:hlinkClick r:id=""/>`` is the ordinary encoding of an action-only
    hyperlink — every "go to next slide" action button is one — and a real
    deck is full of them.

    Two mechanical rules, both load-bearing:

    1. **Attributes are read before any are written.** Aliasing is real: a map
       containing both ``rId1 -> rId2`` *and* ``rId2 -> rId1`` is a legal
       outcome of a collapse plus a gap-fill, and any pass that re-reads an
       attribute it has already written applies one of them twice. This is
       also why the rewrite must never be a string substitution over the
       serialized blob, which is the shortcut it is tempting to reach for.
    2. **Run it per cloned XML part, not once over the slide.** A chart's
       ``c:externalData/@r:id`` and a diagram's ``@relId`` live in sub-parts,
       each with its own relationships and therefore its own map.

    Example:
        >>> from pptx_plus.core.ns import qn
        >>> from pptx_plus.core.oxml import el
        >>> node = el("a:blip", **{"r:embed": "rId3"})
        >>> remap_rel_ids(node, {"rId3": "rId2"})
        1
        >>> node.get(qn("r:embed"))
        'rId2'
    """
    edits: list[tuple[etree._Element, str, str]] = []
    for node, name, value in _rel_id_attrs(root):
        if not value:
            continue
        replacement = relmap.get(value)
        if replacement is None:
            if strict:
                raise DanglingRelationshipError(
                    f"<{_local(node)} {_local_attr(name)}={value!r}> names a relationship "
                    f"that does not exist on the source part. Known ids: "
                    f"{sorted(relmap)}"
                )
            continue
        edits.append((node, name, replacement))

    for node, name, replacement in edits:
        node.set(name, replacement)

    _warn_unclaimed(root)
    return len(edits)


def _warn_unclaimed(root: etree._Element) -> None:
    """Emit :class:`RelIdLiteralWarning` for rel-id-shaped values nobody claims."""
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue
        for key, value in node.attrib.items():
            name = str(key)
            if name.startswith(_R_PREFIX) or (node.tag, name) in UNQUALIFIED_REL_ID_ATTRS:
                continue
            if _RID_LITERAL.match(str(value)):
                warnings.warn(
                    f"<{_local(node)} {_local_attr(name)}={str(value)!r}> holds a value "
                    f"shaped like a relationship id but is not a known "
                    f"relationship-id attribute, so it was not rewritten. If it is one, "
                    f"add it to UNQUALIFIED_REL_ID_ATTRS in pptx_plus.core.reltypes.",
                    RelIdLiteralWarning,
                    stacklevel=2,
                )


def _local(node: etree._Element) -> str:
    """Return an element's local name, for a readable message."""
    tag = str(node.tag)
    return tag.rpartition("}")[2] if "}" in tag else tag


def _local_attr(name: str) -> str:
    """Return an attribute's name in ``prefix:local`` form where recognizable."""
    if name.startswith(_R_PREFIX):
        return f"r:{name[len(_R_PREFIX) :]}"
    return name.rpartition("}")[2] if "}" in name else name


__all__ = [
    "DanglingRelationshipError",
    "RelIdLiteralWarning",
    "RelMap",
    "remap_rel_ids",
]
