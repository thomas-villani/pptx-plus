"""Upstream-surface guard for python-pptx.

``pptx_plus`` reaches below python-pptx's public object model to the OPC
layer. Most of what it uses there is public and stable — ``Part.load``,
``Part.blob``, ``Part.relate_to``, ``Part.rels``, ``Package.next_partname`` —
but "most" is not "all", and none of it is covered by python-pptx's own
compatibility promises.

Rather than discover a breaking change as a mysterious failure in a user's
deck six months after an upgrade, :func:`check_upstream_surface` asserts every
depended-on attribute exists at import time and names the missing one if not.
``tests/test_upstream_surface.py`` covers each individually, so a version bump
produces one red test naming the attribute on the day it is bumped.

This is the pptx analogue of docx-plus's ``tests/test_import_invariant.py``:
turn a convention that would otherwise drift into a check that cannot.

SPEC §14.2.
"""

from __future__ import annotations

from pptx_plus.core.errors import PptxPlusError

#: The python-pptx releases this surface has actually been verified against.
#: Quoted in the error message so a report says which versions were tested,
#: rather than leaving a user to guess.
TESTED_VERSIONS = ">=1.0.2,<2"


class UpstreamSurfaceError(PptxPlusError, RuntimeError):
    """Raised when python-pptx no longer provides an attribute pptx_plus needs.

    Subclasses ``RuntimeError`` because it signals an environment problem
    rather than a caller mistake; also subclasses
    :class:`~pptx_plus.core.errors.PptxPlusError` per SPEC §9.7.
    """


#: ``(import path, attribute, why we need it)``, checked in order.
#:
#: Keep the third element concrete. Its whole job is to tell whoever reads the
#: failure what breaks and what to look for upstream — "needed by the clone
#: engine" is useless, "the byte-faithful part clone primitive" is not.
REQUIRED_SURFACE: tuple[tuple[str, str, str], ...] = (
    # --- The clone primitive. Public, and no part class overrides `load`, so
    # --- one expression clones slides, charts, images, and blob-only parts
    # --- alike. This is what keeps `_element` out of the clone path entirely.
    ("pptx.opc.package", "Part.load", "the byte-faithful part clone primitive"),
    ("pptx.opc.package", "Part.blob", "the source bytes for a part clone"),
    ("pptx.opc.package", "Part.partname", "part identity"),
    ("pptx.opc.package", "Part.content_type", "the content type carried onto a clone"),
    ("pptx.opc.package", "Part.package", "reaching the package from a part"),
    ("pptx.opc.package", "Part.relate_to", "minting a relationship on a cloned part"),
    ("pptx.opc.package", "Part.rels", "enumerating a part's relationships"),
    # --- Package-level allocation and traversal.
    ("pptx.opc.package", "OpcPackage.next_partname", "allocating a free part name"),
    ("pptx.opc.package", "OpcPackage.iter_parts", "relationship-graph reachability"),
    # --- The relationship collection. `pop` is unconditional; XmlPart.drop_rel
    # --- is not, and silently no-ops on a slide referenced from a custom show.
    ("pptx.opc.package", "_Relationships.pop", "unconditional relationship removal"),
    ("pptx.opc.package", "_Relationships.get_or_add", "re-minting an internal relationship"),
    (
        "pptx.opc.package",
        "_Relationships.get_or_add_ext_rel",
        "re-minting an external relationship",
    ),
    ("pptx.opc.package", "_Relationship.rId", "the source key of a relationship-id remap"),
    ("pptx.opc.package", "_Relationship.reltype", "share-vs-deep clone classification"),
    (
        "pptx.opc.package",
        "_Relationship.is_external",
        "external relationships are re-minted by ref",
    ),
    ("pptx.opc.package", "_Relationship.target_part", "walking to a related part"),
    ("pptx.opc.package", "_Relationship.target_ref", "the target of an external relationship"),
    # --- Slide-list manipulation. The one genuinely private area, quarantined
    # --- in `core/oxml.py`'s `sld_id_lst()`.
    ("pptx.oxml.presentation", "CT_Presentation.get_or_add_sldIdLst", "reaching p:sldIdLst"),
    ("pptx.oxml.presentation", "CT_SlideIdList.add_sldId", "appending a slide entry"),
    ("pptx.oxml.presentation", "CT_SlideId.id", "the deck-scoped slide id"),
    ("pptx.oxml.presentation", "CT_SlideId.rId", "the part-scoped relationship id"),
    # --- Slide identity. `resolve_slide` matches a Slide to its position by
    # --- the identity of its `<p:sld>` element, since python-pptx defines no
    # --- `__eq__` on Slide and index is not carried on the object.
    ("pptx.slide", "Slide.element", "locating a slide by element identity"),
    # --- Notes. `has_notes_slide` is the non-mutating test; `notes_slide` is a
    # --- lazyproperty that CREATES parts and must never be used to inspect.
    ("pptx.parts.slide", "SlidePart.has_notes_slide", "testing for notes without creating them"),
)


def _resolve(module_path: str, dotted: str) -> bool:
    """Return whether ``dotted`` resolves inside ``module_path``.

    Attributes are looked up on the class rather than an instance, so nothing
    is constructed and no lazy property can fire as a side effect of checking.
    """
    import importlib

    try:
        obj: object = importlib.import_module(module_path)
    except ImportError:
        return False
    for name in dotted.split("."):
        try:
            obj = getattr(obj, name)
        except AttributeError:
            return False
    return True


def check_upstream_surface() -> None:
    """Verify python-pptx still exposes everything pptx_plus depends on.

    Raises:
        UpstreamSurfaceError: An expected attribute is missing. The message
            names the attribute, what it is needed for, and the version range
            this release was verified against.

    Called once from ``pptx_plus/__init__.py``. The cost is a handful of
    ``getattr`` calls on already-imported modules.
    """
    missing = [
        f"{module_path}.{dotted} ({why})"
        for module_path, dotted, why in REQUIRED_SURFACE
        if not _resolve(module_path, dotted)
    ]
    if missing:
        import pptx

        raise UpstreamSurfaceError(
            f"python-pptx {pptx.__version__} is missing {len(missing)} attribute(s) "
            f"pptx_plus requires. This release was verified against python-pptx "
            f"{TESTED_VERSIONS}. Missing: " + "; ".join(missing)
        )


__all__ = ["REQUIRED_SURFACE", "TESTED_VERSIONS", "UpstreamSurfaceError", "check_upstream_surface"]
