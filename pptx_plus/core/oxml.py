"""Element construction and query — the single chokepoint for OOXML I/O.

Capability modules construct and query elements through this module rather
than calling ``lxml.etree`` directly. SPEC §9.2.

It also owns the *entire* private-API surface this library touches on
python-pptx objects. :func:`part_root` and :func:`sld_id_lst` are the only two
functions anywhere in ``pptx_plus`` that reach for an underscore-prefixed
attribute, and ``tests/test_upstream_surface.py`` covers both. Everything else
— cloning parts, minting relationships, allocating part names — runs on
python-pptx's public API. SPEC §14.2.
"""

from __future__ import annotations

from functools import lru_cache
from typing import TYPE_CHECKING, Any, cast

from lxml import etree

from pptx_plus.core.ns import BUILD_NSMAP, NSMAP, qn

if TYPE_CHECKING:
    from pptx.opc.package import XmlPart
    from pptx.oxml.presentation import CT_Presentation, CT_SlideIdList
    from pptx.presentation import Presentation


def _resolve_attr_key(key: str) -> str:
    """Translate a ``"prefix:local"`` attribute key to Clark notation.

    Plain keys (no ``:``) are returned unchanged, so callers can mix
    namespaced and bare attributes naturally — ``p:sldId`` carries a
    namespaced ``r:id`` next to a bare ``id``, and both are ordinary here.
    """
    return qn(key) if ":" in key else key


def el(tag: str, **attrs: str) -> etree._Element:
    """Create a namespaced element with attributes.

    Args:
        tag: Element name in ``prefix:local`` form. The prefix must be a key
            in :data:`~pptx_plus.core.ns.NSMAP`.
        **attrs: Attributes. Keys may be namespaced (``"r:id"``) or plain
            (``"id"``).

    Returns:
        A fresh detached :class:`lxml.etree._Element`.

    Raises:
        InvalidNamespaceError: The tag or an attribute key names an unknown
            prefix.

    Elements in the main presentation namespaces are built declaring
    :data:`~pptx_plus.core.ns.BUILD_NSMAP`. An element in any other namespace
    declares only its own prefix, so extension parts stay free of irrelevant
    declarations.

    Note:
        Because OOXML attribute names contain a colon, call sites have to
        dict-splat rather than pass keywords directly::

            el("p:sldId", **{"id": "256", "r:id": "rId2"})

        That is ugly, and it is the price of having one place where every
        element in this library is born.

    Example:
        >>> el("p:sldId", id="256").tag.endswith("}sldId")
        True
    """
    qname = qn(tag)  # validates the prefix before it is used as an nsmap key
    prefix = tag.partition(":")[0]
    nsmap = BUILD_NSMAP if prefix in BUILD_NSMAP else {prefix: NSMAP[prefix]}
    node = etree.Element(qname, nsmap=nsmap)
    for key, value in attrs.items():
        node.set(_resolve_attr_key(key), value)
    return node


def sub(parent: etree._Element, tag: str, **attrs: str) -> etree._Element:
    """Create an element via :func:`el` and append it to ``parent``.

    Args:
        parent: The element to append to.
        tag: Element name in ``prefix:local`` form.
        **attrs: Attributes, as for :func:`el`.

    Returns:
        The newly created and appended element.
    """
    child = el(tag, **attrs)
    parent.append(child)
    return child


@lru_cache(maxsize=512)
def _compile_xpath(expr: str) -> etree.XPath:
    """Compile and cache an XPath expression bound to :data:`NSMAP`.

    Expressions are drawn from a small set of literals in this codebase, so
    the cache hit rate is effectively 100% after warmup.
    """
    return etree.XPath(expr, namespaces=NSMAP)


def xpath(node: etree._Element, expr: str, **variables: Any) -> list[Any]:
    """Evaluate an XPath expression against ``node``.

    Args:
        node: The context node.
        expr: An XPath expression. Prefixes resolve against
            :data:`~pptx_plus.core.ns.NSMAP`.
        **variables: Values bound to ``$name`` references in ``expr``.

    Returns:
        The result list.

    Pass values as **XPath variables**, never by interpolating them into the
    expression::

        xpath(root, "./p:sldId[@r:id=$rid]", rid=r_id)     # yes
        xpath(root, f"./p:sldId[@r:id='{r_id}']")          # no

    Interpolation breaks on any value containing a quote and defeats the
    compiled-expression cache, since every distinct value produces a distinct
    expression.

    Note:
        python-pptx's ``BaseOxmlElement.xpath()`` does not accept a
        ``namespaces=`` keyword and resolves prefixes against python-pptx's own
        map, which lacks ``mc``, ``dgm``, ``dsp``, and the extension
        namespaces. Use this function rather than the method.
    """
    result = _compile_xpath(expr)(node, **variables)
    return list(result) if isinstance(result, list) else [result]


def remove(node: etree._Element) -> None:
    """Detach ``node`` from its parent.

    A no-op if the node is already detached, which makes the delete paths
    that call it re-entrant rather than order-sensitive.
    """
    parent = node.getparent()
    if parent is not None:
        parent.remove(node)


def ordered_insert(
    parent: etree._Element,
    child: etree._Element,
    order: tuple[str, ...],
) -> etree._Element:
    """Insert ``child`` into ``parent`` at its schema-mandated position.

    Args:
        parent: The parent element.
        child: The element to insert.
        order: ``parent``'s full child sequence in schema order, as
            ``prefix:local`` names.

    Returns:
        The inserted child.

    Raises:
        ValueError: ``child``'s tag is not in ``order``.

    Any existing sibling with the same tag is removed first, which makes the
    operation **idempotent**: calling it twice leaves one element, not two.
    """
    tag = etree.QName(child).text
    names = [qn(name) for name in order]
    if tag not in names:
        raise ValueError(f"ordered_insert: {tag!r} is not in the declared child order {order!r}")

    for existing in list(parent):
        if existing.tag == tag:
            parent.remove(existing)

    successors = set(names[names.index(tag) + 1 :])
    for sibling in parent:
        if isinstance(sibling.tag, str) and sibling.tag in successors:
            sibling.addprevious(child)
            return child
    parent.append(child)
    return child


# ---------------------------------------------------------------------------
# The private-API quarantine.
#
# Everything below reaches into python-pptx internals. Keeping it to two
# functions in one module means an upstream change breaks two call sites, and
# `tests/test_upstream_surface.py` says which. SPEC §14.2.
# ---------------------------------------------------------------------------


def part_root(part: XmlPart) -> etree._Element:
    """Return the root element of an XML part.

    python-pptx exposes a part's parsed tree only as ``XmlPart._element``.
    There is no public accessor, and ``part.blob`` re-serializes rather than
    handing back the live tree, so it is not a substitute when the tree is
    about to be mutated.

    Args:
        part: Any ``XmlPart`` — a slide, the presentation part, a chart.

    Returns:
        The live root element. Mutating it mutates the part.
    """
    return part._element  # noqa: SLF001


def sld_id_lst(prs: Presentation) -> CT_SlideIdList:
    """Return the presentation's ``<p:sldIdLst>``, creating it if absent.

    This element is the **sole** determinant of slide order (SPEC §3.4), so
    every verb in ``slides/`` goes through here rather than touching
    ``prs.slides._sldIdLst`` — the entry point of the naive recipes — directly.

    Args:
        prs: The presentation.

    Returns:
        The ``<p:sldIdLst>`` element.

    Uses python-pptx's ``CT_Presentation.get_or_add_sldIdLst()``, which places
    a newly created element correctly relative to its schema successors
    (``p:sldSz``, ``p:notesSz``). A deck with no slides may genuinely lack the
    element, so creating it is the right behaviour rather than an error.
    """
    presentation = cast("CT_Presentation", part_root(prs.part))
    return presentation.get_or_add_sldIdLst()


__all__ = [
    "el",
    "ordered_insert",
    "part_root",
    "remove",
    "sld_id_lst",
    "sub",
    "xpath",
]
