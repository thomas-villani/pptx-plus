"""Namespace map and qualified-name helper — SPEC §4.1."""

from __future__ import annotations

import pytest

from pptx_plus.core.ns import BUILD_NSMAP, NSMAP, InvalidNamespaceError, qn


def test_qn_expands_a_known_prefix() -> None:
    assert qn("p:sldId") == f"{{{NSMAP['p']}}}sldId"


def test_qn_rejects_an_unqualified_name() -> None:
    with pytest.raises(InvalidNamespaceError, match="expected 'prefix:local'"):
        qn("sldId")


def test_qn_rejects_an_unknown_prefix() -> None:
    with pytest.raises(InvalidNamespaceError, match="unknown namespace prefix 'nope'"):
        qn("nope:thing")


def test_invalid_namespace_error_is_a_value_error() -> None:
    """SPEC §16: typed errors dual-inherit the stdlib type a caller would catch."""
    assert issubclass(InvalidNamespaceError, ValueError)


def test_qn_failures_are_not_cached() -> None:
    """`@cache` stores return values; a raised exception leaves no entry.

    Worth pinning: a cached failure would turn a transient into a permanent.
    """
    for _ in range(2):
        with pytest.raises(InvalidNamespaceError):
            qn("still-not-valid")


@pytest.mark.parametrize("prefix", sorted(BUILD_NSMAP))
def test_build_nsmap_is_a_subset_of_nsmap(prefix: str) -> None:
    """The write map must never bind a prefix the query map disagrees with."""
    assert BUILD_NSMAP[prefix] == NSMAP[prefix]


def test_build_nsmap_is_small() -> None:
    """v0.1 is a rewriter, not a builder — SPEC §4.1.

    This is a design tripwire rather than a correctness check. Growth here
    means a capability module has started authoring element trees, which is
    the point at which python-pptx's own model is the better tool. If this
    fails, the question to answer is "should this code be building XML at
    all?", not "should this number be bigger?".
    """
    assert set(BUILD_NSMAP) == {"p", "a", "r"}


@pytest.mark.parametrize(
    ("prefix", "expected"),
    [
        # The prefixes python-pptx's own map omits, which is the entire reason
        # this module exists rather than delegating.
        ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
        ("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram"),
        ("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main"),
        ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
        ("a16", "http://schemas.microsoft.com/office/drawing/2014/main"),
        ("asvg", "http://schemas.microsoft.com/office/drawing/2016/SVG/main"),
    ],
)
def test_extension_prefixes_resolve(prefix: str, expected: str) -> None:
    assert NSMAP[prefix] == expected


def test_python_pptx_cannot_resolve_mc() -> None:
    """The concrete justification for owning our own map.

    python-pptx binds the markup-compatibility URI to `ve`, not `mc`, so
    `pptx.oxml.ns.qn("mc:AlternateContent")` raises. Everything interesting
    that a part-level clone preserves — embedded media, SVG blips, modern
    image effects — lives under `mc:AlternateContent`.

    If this ever stops raising, upstream has added the prefix and this module
    could shrink. That is worth being told about.
    """
    from pptx.oxml.ns import qn as pptx_qn

    with pytest.raises(KeyError):
        pptx_qn("mc:AlternateContent")
