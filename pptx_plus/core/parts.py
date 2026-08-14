"""Part-name allocation, byte-faithful part cloning, relationship removal.

Three primitives the clone engine and the delete path are built from. All
three are small, and all three exist because the obvious version is wrong in a
way that fails silently.

SPEC §4.6. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pptx.opc.packuri import PackURI

from pptx_plus.core.errors import PptxPlusError
from pptx_plus.core.reltypes import PARTNAME_TEMPLATES

if TYPE_CHECKING:
    from pptx.opc.package import Part

    # `pptx.package.Package` is the concrete subclass of `OpcPackage` that
    # `Part.load` is annotated against, and the only one a real deck produces.
    from pptx.package import Package

#: Matches a part name ending in digits before its extension, so a template can
#: be derived from an existing name when the content type is unrecognized.
_NUMBERED_PARTNAME = re.compile(r"^(?P<stem>.*?)(?P<index>\d+)(?P<ext>\.[^.]+)$")


class UnclonablePartError(PptxPlusError, TypeError):
    """Raised when a part cannot be cloned.

    Subclasses ``TypeError`` because it signals an unsupported input type;
    also subclasses :class:`~pptx_plus.core.errors.PptxPlusError` per
    SPEC §9.7.
    """


def partname_template_for(part: Part) -> str:
    """Return a ``next_partname`` template appropriate for ``part``.

    Args:
        part: The source part.

    Returns:
        A template such as ``"/ppt/slides/slide%d.xml"``.

    Raises:
        UnclonablePartError: The part's name has no numeric component and its
            content type is unknown, so no template can be derived.

    Known content types come from :data:`~pptx_plus.core.reltypes.PARTNAME_TEMPLATES`.
    Anything else has a template derived from its *own* name, so a part type
    this library has never seen still clones — landing beside its original,
    which is both correct and what someone unzipping the result would expect.
    Refusing to clone unknown parts would defeat the whole preservation
    guarantee, since "unknown to us" describes most of what makes a real deck
    interesting.
    """
    template = PARTNAME_TEMPLATES.get(part.content_type)
    if template is not None:
        return template

    match = _NUMBERED_PARTNAME.match(str(part.partname))
    if match is None:
        raise UnclonablePartError(
            f"cannot derive a part-name template for {part.partname!r} "
            f"(content type {part.content_type!r}): the name has no numeric "
            f"component and the content type is not in PARTNAME_TEMPLATES"
        )
    return f"{match['stem']}%d{match['ext']}"


def allocate_partname(
    package: Package,
    template: str,
    reserved: set[str],
) -> PackURI:
    """Allocate an unused part name, honouring an in-flight reservation set.

    Args:
        package: The target package.
        template: A ``printf``-style template with a single ``%d``.
        reserved: Names already handed out during this operation. **Mutated**:
            the returned name is added before returning.

    Returns:
        A :class:`~pptx.opc.packuri.PackURI` unused in both the package and
        ``reserved``.

    ``reserved`` is not an optimization and not optional.
    ``OpcPackage.next_partname`` derives its used-name set from
    ``iter_parts()``, which walks the relationship graph from the package
    root — so **a part that has been constructed but not yet related to
    anything is invisible to it.**

    Cloning a slide bearing two charts hits this immediately: the first clone
    exists but is not yet attached when the second is allocated, so both are
    handed ``/ppt/charts/chart2.xml`` and the second zip entry silently
    overwrites the first. Attaching before recursing is not an option, because
    the parent's relationship does not exist yet — hence the reservation set.
    """
    index = 1
    while True:
        candidate = template % index
        if candidate not in reserved and not _in_package(package, candidate):
            reserved.add(candidate)
            return PackURI(candidate)
        index += 1


def _in_package(package: Package, partname: str) -> bool:
    """Return whether ``partname`` is already reachable in ``package``."""
    return any(str(part.partname) == partname for part in package.iter_parts())


def clone_part(
    src: Part,
    *,
    into: Package,
    reserved: set[str],
    blob: bytes | None = None,
) -> Part:
    """Clone ``src`` byte-faithfully into ``into`` under a fresh part name.

    Args:
        src: The part to clone.
        into: The destination package.
        reserved: In-flight part-name reservations; see :func:`allocate_partname`.
        blob: Content for the clone, replacing the source's own bytes. The one
            caller is the clone engine, for a part whose *contents* hold
            relationship ids scoped to the referring part — a SmartArt data
            part — which therefore has to be rewritten before it is loaded
            rather than after. Everything else takes the default and stays
            byte-identical. SPEC §4.5.

    Returns:
        A new part of the same class and content type. It has **no
        relationships** — minting those is the caller's job, because the ids
        they produce are what the XML rewrite needs.

    The whole operation is one expression::

        type(src).load(partname, src.content_type, into, src.blob)

    ``Part.load`` and ``Part.blob`` are both public, and no part class in
    python-pptx overrides ``load`` — asserted by
    ``tests/test_upstream_surface.py``, because the day one does, this stops
    being uniform. That uniformity is what lets slides, charts, images,
    embedded workbooks, and every blob-only SmartArt part take the same path.

    For a blob-backed part — every diagram definition part, every media file,
    every embedded package — ``blob`` is the original byte string and ``load``
    stores it unchanged. No reparse, so no whitespace normalization and no
    re-serialization. That is what makes the preservation guarantee in
    SPEC §8.1 byte-exact rather than merely close.
    """
    partname = allocate_partname(into, partname_template_for(src), reserved)
    return type(src).load(
        partname,
        src.content_type,
        into,
        src.blob if blob is None else blob,
    )


def drop_relationship(part: Part, rId: str) -> None:  # noqa: N803 - OOXML spelling
    """Remove relationship ``rId`` from ``part`` unconditionally.

    Args:
        part: The part holding the relationship.
        rId: The relationship id to remove.

    Raises:
        KeyError: No such relationship on this part.

    **Do not replace this with ``XmlPart.drop_rel``.** That method is
    conditional — it removes the relationship only when its reference count in
    the part's XML is below 2, counting matches of ``//@r:id``. On the
    presentation part, a slide that also appears in a custom show is
    referenced twice, so ``drop_rel`` silently does nothing and leaves the
    slide serialized but unreachable from ``sldIdLst``.

    ``_Relationships.pop`` is unconditional, and ``Part.rels`` is documented
    upstream as necessarily public so the part graph can be traversed.
    """
    part.rels.pop(rId)


__all__ = [
    "UnclonablePartError",
    "allocate_partname",
    "clone_part",
    "drop_relationship",
    "partname_template_for",
]
