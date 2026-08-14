"""XML namespace map and qualified-name helper.

python-pptx has its own map in ``pptx.oxml.ns``, and it is not sufficient
here. It omits every prefix the copy engine has to be able to *query*: there
is no ``mc`` (it binds the markup-compatibility URI to ``ve``), no ``dgm`` or
``dsp`` for SmartArt, and none of the Microsoft extension namespaces —
``p14``, ``a14``, ``a16``, ``asvg`` — under which embedded media and modern
image effects live. ``pptx.oxml.ns.qn("mc:AlternateContent")`` raises
``KeyError``. So this module owns the map.

Two maps, for two different jobs:

- :data:`NSMAP` is the **query** map. Every prefix the library can address,
  used to compile XPath and to resolve :func:`qn`.
- :data:`BUILD_NSMAP` is the **write** map, declared on elements this library
  constructs. It is deliberately tiny — see its docstring.

SPEC §4.1.
"""

from __future__ import annotations

from functools import cache

from pptx_plus.core.errors import PptxPlusError

# --- ECMA-376 core namespaces ---------------------------------------------

#: PresentationML — ``p:sld``, ``p:sldIdLst``, ``p:custShowLst``.
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
#: DrawingML main — ``a:blip``, ``a:hlinkClick``, ``a:ext``.
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
#: Relationship references. **Every attribute in this namespace is a
#: relationship id** — see :mod:`pptx_plus.core.relmap`, which relies on that.
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#: DrawingML charts — ``c:chartSpace``, ``c:externalData``.
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
#: DrawingML pictures — ``pic:pic``, ``pic:blipFill``.
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
#: DrawingML diagrams (SmartArt) — the data model, layout, style, and colour
#: definition parts.
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
#: Markup Compatibility — ``mc:AlternateContent``, ``mc:Choice``,
#: ``mc:Fallback``. python-pptx binds this URI to ``ve``, not ``mc``.
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# --- OPC package namespaces -----------------------------------------------
#
# Needed by the integrity harness, which reads `.rels` files and
# `[Content_Types].xml` straight out of the saved zip rather than through
# python-pptx — precisely so it grades the artifact rather than the model.

#: Package relationships — the root element of every ``_rels/*.rels`` part.
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
#: Content types — the root element of ``[Content_Types].xml``.
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# --- Microsoft extension namespaces ---------------------------------------
#
# These carry content python-pptx does not model, and they are exactly what a
# part-level clone preserves and a model round-trip would strip. They appear
# almost exclusively inside `mc:AlternateContent`.

#: SmartArt rendered drawing cache — ``dsp:drawing``, and the
#: ``dsp:dataModelExt/@relId`` pointer, which is the one relationship id in a
#: real deck that is *not* in the ``r:`` namespace.
DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
#: PowerPoint 2010 extensions — ``p14:media``, ``p14:sectionLst``.
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
#: Drawing 2010 extensions — ``a14:imgLayer`` and the image-effect stack.
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
#: Drawing 2014 extensions — ``a16:creationId``.
A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
#: SVG blips — ``asvg:svgBlip``, the vector companion to a raster ``a:blip``.
ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"

# --- Legacy VML / Office namespaces ---------------------------------------

#: Legacy Office drawing — ``o:OLEObject``.
O = "urn:schemas-microsoft-com:office:office"  # noqa: E741
#: VML — ``v:imagedata``, still used for OLE object fallback rendering.
V = "urn:schemas-microsoft-com:vml"

#: The XML namespace itself — ``xml:lang``, ``xml:space``.
XML = "http://www.w3.org/XML/1998/namespace"


#: The **query** map: every prefix this library can address. Used to compile
#: XPath expressions and to resolve :func:`qn`. Adding a prefix here is cheap
#: and has no effect on output.
NSMAP: dict[str, str] = {
    "p": P,
    "a": A,
    "r": R,
    "c": C,
    "pic": PIC,
    "dgm": DGM,
    "mc": MC,
    "pr": PR,
    "ct": CT,
    "dsp": DSP,
    "p14": P14,
    "a14": A14,
    "a16": A16,
    "asvg": ASVG,
    "o": O,
    "v": V,
    "xml": XML,
}

#: The **write** map, declared on elements built by :func:`~pptx_plus.core.oxml.el`.
#:
#: Deliberately tiny, and that is the structural difference between this
#: library and its docx-plus sibling. docx-plus is a *builder*: it authors
#: whole element trees, so its write map has to be broad. pptx_plus v0.1 is a
#: *rewriter* — it clones existing XML and edits attribute values in place.
#: The only element it constructs is ``p:sldId``, and python-pptx's
#: ``CT_SlideIdList.add_sldId`` constructs even that.
#:
#: Treat growth in this map as a signal worth questioning: it means a
#: capability module has started authoring rather than cloning, which is the
#: point at which python-pptx's own model becomes the better tool. SPEC §4.1.
BUILD_NSMAP: dict[str, str] = {"p": P, "a": A, "r": R}


class InvalidNamespaceError(PptxPlusError, ValueError):
    """Raised for a malformed or unknown ``prefix:local`` name.

    Subclasses ``ValueError`` so existing ``except ValueError:`` clauses still
    catch it; also subclasses :class:`~pptx_plus.core.errors.PptxPlusError`
    per SPEC §9.7.
    """


@cache
def qn(name: str) -> str:
    """Translate a ``"prefix:local"`` name to Clark notation.

    Args:
        name: A qualified name such as ``"p:sldId"``. The prefix must be a key
            in :data:`NSMAP`.

    Returns:
        The name in lxml's Clark notation, e.g. ``"{http://…/main}sldId"``.

    Raises:
        InvalidNamespaceError: The name has no colon, or its prefix is unknown.

    Memoized because it is called once per element and once per attribute on
    every pass over a part — a single slide clone runs it thousands of times,
    and the inputs are drawn from a set of a few dozen literals. Failures are
    not cached: ``@cache`` stores return values, and a raised exception
    propagates without an entry being made.

    Example:
        >>> qn("p:sldId").endswith("}sldId")
        True
    """
    prefix, sep, local = name.partition(":")
    if not sep:
        raise InvalidNamespaceError(f"qn() expected 'prefix:local', got {name!r}")
    try:
        uri = NSMAP[prefix]
    except KeyError as exc:
        raise InvalidNamespaceError(f"unknown namespace prefix {prefix!r} in {name!r}") from exc
    return f"{{{uri}}}{local}"


__all__ = [
    "A",
    "A14",
    "A16",
    "ASVG",
    "BUILD_NSMAP",
    "C",
    "CT",
    "DGM",
    "DSP",
    "MC",
    "NSMAP",
    "O",
    "P",
    "P14",
    "PIC",
    "PR",
    "R",
    "V",
    "XML",
    "InvalidNamespaceError",
    "qn",
]
