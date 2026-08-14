"""Rewriting relationship ids -- SPEC §4.4.

The step every circulating "duplicate a slide" recipe omits, tested standalone
before anything is wired to it.
"""

from __future__ import annotations

import pytest

from pptx_plus.core.ns import qn
from pptx_plus.core.oxml import el, sub
from pptx_plus.core.relmap import (
    DanglingRelationshipError,
    RelIdLiteralWarning,
    remap_rel_ids,
)

# ---------------------------------------------------------------------------
# The `r:` namespace sweep
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("attr", ["r:id", "r:embed", "r:link", "r:dm", "r:lo", "r:qs", "r:cs"])
def test_every_r_namespace_attribute_is_rewritten(attr: str) -> None:
    """The namespace is closed by schema, so the sweep needs no element list."""
    node = el("a:blip", **{attr: "rId3"})
    remap_rel_ids(node, {"rId3": "rId9"})
    assert node.get(qn(attr)) == "rId9"


def test_a_nested_attribute_is_rewritten() -> None:
    root = el("p:sld")
    sub(sub(root, "p:cSld"), "a:blip", **{"r:embed": "rId3"})
    remap_rel_ids(root, {"rId3": "rId9"})
    assert root[0][0].get(qn("r:embed")) == "rId9"


def test_an_element_this_library_has_never_heard_of_is_rewritten() -> None:
    """The argument for a namespace sweep over an element allowlist.

    `asvg:svgBlip` is not in any table here. It is rewritten anyway, because
    what identifies a relationship reference is the attribute's namespace, not
    the element's name -- and element names are an open vocabulary that grows
    with every Office release.
    """
    node = el("asvg:svgBlip", **{"r:embed": "rId3"})
    remap_rel_ids(node, {"rId3": "rId9"})
    assert node.get(qn("r:embed")) == "rId9"


def test_the_count_of_rewrites_is_returned() -> None:
    root = el("p:sld")
    for _ in range(3):
        sub(root, "a:blip", **{"r:embed": "rId3"})
    assert remap_rel_ids(root, {"rId3": "rId9"}) == 3


def test_nothing_to_rewrite_returns_zero() -> None:
    assert remap_rel_ids(el("p:sld"), {"rId1": "rId2"}) == 0


def test_a_non_relationship_attribute_is_left_alone() -> None:
    node = el("p:cNvPr", id="7", name="Picture 3")
    remap_rel_ids(node, {"rId1": "rId2"})
    assert node.get("name") == "Picture 3"


# ---------------------------------------------------------------------------
# Aliasing -- the reason this is not a string substitution
# ---------------------------------------------------------------------------


def test_a_swap_is_applied_once_not_twice() -> None:
    """The flagship case. SPEC §4.4.

    A map holding both `rId1 -> rId2` and `rId2 -> rId1` is a legal outcome of
    a relationship collapse plus a gap-fill. Any pass that re-reads an
    attribute it has already written applies one of the two mappings twice --
    which is exactly what a regex over the serialized XML does, and a regex is
    the shortcut this failing test exists to forbid.
    """
    root = el("p:sld")
    first = sub(root, "a:blip", **{"r:embed": "rId1"})
    second = sub(root, "a:blip", **{"r:embed": "rId2"})

    remap_rel_ids(root, {"rId1": "rId2", "rId2": "rId1"})

    assert (first.get(qn("r:embed")), second.get(qn("r:embed"))) == ("rId2", "rId1")


def test_a_chained_mapping_is_not_followed() -> None:
    """`rId1 -> rId2 -> rId3` rewrites rId1 to rId2, not to rId3."""
    node = el("a:blip", **{"r:embed": "rId1"})
    remap_rel_ids(node, {"rId1": "rId2", "rId2": "rId3"})
    assert node.get(qn("r:embed")) == "rId2"


def test_a_many_to_one_map_is_accepted() -> None:
    """`get_or_add` dedupes on (reltype, target), so the map is not a bijection."""
    root = el("p:sld")
    first = sub(root, "a:blip", **{"r:embed": "rId1"})
    second = sub(root, "a:blip", **{"r:embed": "rId2"})
    remap_rel_ids(root, {"rId1": "rId1", "rId2": "rId1"})
    assert (first.get(qn("r:embed")), second.get(qn("r:embed"))) == ("rId1", "rId1")


# ---------------------------------------------------------------------------
# The empty id
# ---------------------------------------------------------------------------


def test_an_empty_id_is_not_rewritten() -> None:
    """`<a:hlinkClick r:id=""/>` is an action-only link, not a dangling one."""
    node = el("a:hlinkClick", **{"r:id": "", "action": "ppaction://hlinkshowjump"})
    remap_rel_ids(node, {"rId1": "rId2"})
    assert node.get(qn("r:id")) == ""


def test_an_empty_id_does_not_raise_under_strict() -> None:
    node = el("a:hlinkClick", **{"r:id": ""})
    assert remap_rel_ids(node, {}, strict=True) == 0


# ---------------------------------------------------------------------------
# Unmapped ids
# ---------------------------------------------------------------------------


def test_an_unmapped_id_raises_under_strict() -> None:
    node = el("a:blip", **{"r:embed": "rId7"})
    with pytest.raises(DanglingRelationshipError):
        remap_rel_ids(node, {"rId1": "rId2"})


def test_the_dangling_error_names_the_attribute_and_value() -> None:
    """A failure that does not say which attribute cannot be acted on."""
    node = el("a:blip", **{"r:embed": "rId7"})
    with pytest.raises(DanglingRelationshipError, match=r"blip r:embed='rId7'"):
        remap_rel_ids(node, {"rId1": "rId2"})


def test_the_dangling_error_lists_the_known_ids(deck) -> None:
    node = el("a:blip", **{"r:embed": "rId7"})
    with pytest.raises(DanglingRelationshipError, match=r"\['rId1'\]"):
        remap_rel_ids(node, {"rId1": "rId2"})


def test_an_unmapped_id_is_left_alone_when_not_strict() -> None:
    node = el("a:blip", **{"r:embed": "rId7"})
    assert remap_rel_ids(node, {"rId1": "rId2"}, strict=False) == 0
    assert node.get(qn("r:embed")) == "rId7"


def test_a_dangling_error_is_a_value_error() -> None:
    assert issubclass(DanglingRelationshipError, ValueError)


def test_nothing_is_written_when_a_later_attribute_dangles() -> None:
    """Attributes are collected before any is written, so a raise is atomic."""
    root = el("p:sld")
    good = sub(root, "a:blip", **{"r:embed": "rId1"})
    sub(root, "a:blip", **{"r:embed": "rId7"})
    with pytest.raises(DanglingRelationshipError):
        remap_rel_ids(root, {"rId1": "rId2"})
    assert good.get(qn("r:embed")) == "rId1"


# ---------------------------------------------------------------------------
# The registered exceptions to the sweep
# ---------------------------------------------------------------------------


def test_the_unqualified_smartart_rel_id_is_rewritten() -> None:
    """`dsp:dataModelExt/@relId` -- the one real rel-id outside `r:`.

    Miss it and PowerPoint silently recovers by recomputing the drawing on
    open, while every other renderer draws nothing at all: a failure invisible
    in the one viewer most likely to be used to check.
    """
    node = el("dsp:dataModelExt", relId="rId5")
    remap_rel_ids(node, {"rId5": "rId2"})
    assert node.get("relId") == "rId2"


def test_a_registered_attribute_does_not_warn() -> None:
    node = el("dsp:dataModelExt", relId="rId5")
    remap_rel_ids(node, {"rId5": "rId2"})  # would error via filterwarnings


def test_a_bare_relid_on_another_element_is_not_rewritten() -> None:
    """The registry is keyed on (element, attribute), not on the name alone."""
    node = el("a:blip", relId="rId5")
    with pytest.warns(RelIdLiteralWarning):
        remap_rel_ids(node, {"rId5": "rId2"})
    assert node.get("relId") == "rId5"


# ---------------------------------------------------------------------------
# The unclaimed-literal tripwire
# ---------------------------------------------------------------------------


def test_an_unknown_attribute_holding_a_rel_id_warns() -> None:
    """How the next unqualified relationship attribute gets found."""
    node = el("p:sld", name="rId7")
    with pytest.warns(RelIdLiteralWarning, match="shaped like a relationship id"):
        remap_rel_ids(node, {})


def test_the_warning_says_where_to_register_it() -> None:
    node = el("p:sld", name="rId7")
    with pytest.warns(RelIdLiteralWarning, match="UNQUALIFIED_REL_ID_ATTRS"):
        remap_rel_ids(node, {})


def test_a_value_merely_containing_rid_does_not_warn() -> None:
    """Anchored: "Grid7" and "rId7 and more" are not relationship ids."""
    node = el("p:sld", name="Grid7", descr="see rId7 for details")
    remap_rel_ids(node, {})


def test_the_warning_is_fatal_in_this_suite() -> None:
    """Guard the guard: the tripwire only works if the filter is registered.

    `pytest.warns` passes whether or not the filter exists, so the tests above
    would keep passing if the promotion to an error were lost. This one fails
    the moment it is -- see `pytest_configure` in `tests/conftest.py` for why
    it cannot live in `pyproject.toml`.
    """
    with pytest.raises(RelIdLiteralWarning):
        remap_rel_ids(el("p:sld", name="rId7"), {})


def test_a_comment_in_the_xml_is_skipped() -> None:
    """A real deck can carry XML comments; `node.tag` is not a string then."""
    from lxml import etree

    root = el("p:sld")
    root.append(etree.Comment(" authored by something else "))
    sub(root, "a:blip", **{"r:embed": "rId1"})
    assert remap_rel_ids(root, {"rId1": "rId2"}) == 1


def test_a_processing_instruction_is_skipped() -> None:
    from lxml import etree

    root = el("p:sld")
    root.append(etree.ProcessingInstruction("target", "data"))
    assert remap_rel_ids(root, {}) == 0


def test_the_warning_is_not_a_pptx_plus_error() -> None:
    """Filtered by class; catchable as PptxPlusError would defeat the point."""
    from pptx_plus.core.errors import PptxPlusError

    assert not issubclass(RelIdLiteralWarning, PptxPlusError)
    assert issubclass(RelIdLiteralWarning, UserWarning)
