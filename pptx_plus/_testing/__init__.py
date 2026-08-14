"""Test-only OPC assertions, shipped in the wheel.

Not part of the public API and not covered by the compatibility promise, but
installed rather than left in ``tests/`` for two reasons: downstream projects
that wrap this library can assert the same invariants against their own decks,
and the assertions travel with the code whose output they grade.

Omitted from coverage and exempt from docstring linting (`pyproject.toml`).
SPEC §10.5.
"""

from __future__ import annotations

from pptx_plus._testing.ooxml_asserts import (
    Rel,
    SavedPackage,
    assert_all_xml_parses,
    assert_content_types_complete,
    assert_in_package,
    assert_no_unclaimed_rid_literals,
    assert_not_in_package,
    assert_package_integrity,
    assert_partnames_unique,
    assert_parts_disjoint,
    assert_parts_shared,
    assert_rel_ids_resolve,
    assert_sections_consistent,
    assert_slide_ids_valid,
    assert_slide_rels_consistent,
    related_partnames,
    roundtrip,
    saved,
)

__all__ = [
    "Rel",
    "SavedPackage",
    "assert_all_xml_parses",
    "assert_content_types_complete",
    "assert_in_package",
    "assert_no_unclaimed_rid_literals",
    "assert_not_in_package",
    "assert_package_integrity",
    "assert_partnames_unique",
    "assert_parts_disjoint",
    "assert_parts_shared",
    "assert_rel_ids_resolve",
    "assert_sections_consistent",
    "assert_slide_ids_valid",
    "assert_slide_rels_consistent",
    "related_partnames",
    "roundtrip",
    "saved",
]
