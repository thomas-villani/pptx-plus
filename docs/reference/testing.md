# `pptx_plus._testing`

The OPC integrity battery. Every assertion reads the **saved zip** rather than
a live `Presentation`, because the in-memory object graph keeps stale
references after a delete by design — so an assertion made against the model
grades the wrong artifact.

Test-only and not covered by the compatibility promise, but shipped in the
wheel so a project wrapping this library can assert the same invariants
against its own decks.

## Reading a package

::: pptx_plus._testing.ooxml_asserts.saved

::: pptx_plus._testing.ooxml_asserts.roundtrip

::: pptx_plus._testing.ooxml_asserts.SavedPackage

## The battery

::: pptx_plus._testing.ooxml_asserts.assert_package_integrity

::: pptx_plus._testing.ooxml_asserts.assert_slide_rels_consistent

::: pptx_plus._testing.ooxml_asserts.assert_slide_ids_valid

::: pptx_plus._testing.ooxml_asserts.assert_rel_ids_resolve

::: pptx_plus._testing.ooxml_asserts.assert_no_unclaimed_rid_literals

::: pptx_plus._testing.ooxml_asserts.assert_partnames_unique

::: pptx_plus._testing.ooxml_asserts.assert_sections_consistent

::: pptx_plus._testing.ooxml_asserts.assert_content_types_complete

::: pptx_plus._testing.ooxml_asserts.assert_all_xml_parses

## Sharing and ownership

::: pptx_plus._testing.ooxml_asserts.assert_parts_shared

::: pptx_plus._testing.ooxml_asserts.assert_parts_disjoint

::: pptx_plus._testing.ooxml_asserts.assert_in_package

::: pptx_plus._testing.ooxml_asserts.assert_not_in_package
