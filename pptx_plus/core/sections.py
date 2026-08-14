"""Section and custom-show maintenance — the side-indexes nobody models.

A deck can index its slides three times over. ``p:sldIdLst`` is the running
order and the only one python-pptx knows about. ``p14:sectionLst``, living in
``p:presentation/p:extLst``, groups slides into the named sections the slide
sorter shows. ``p:custShowLst`` names subsets of slides as alternative
running orders.

python-pptx models neither of the latter two. They survive a round trip only
because unrecognized XML is preserved verbatim — which is exactly what makes
them dangerous: an operation that edits ``sldIdLst`` and stops leaves them
pointing at a slide that is no longer there, and PowerPoint reports the file
as damaged and offers to repair it.

No generated fixture has either structure, so this module's failure mode is
invisible to a test suite that only exercises decks it built itself. It is
covered here by decks assembled by hand and by the committed PowerPoint
samples.

**The two structures key on different identifiers**, which is the trap.
Sections reference ``p:sldId/@id`` — the deck-scoped slide id. Custom shows
reference ``@r:id`` — the part-scoped relationship id on the presentation
part. Same slide, two names, a hundred lines apart in the same file. SPEC §3.3.

ECMA-376 Part 1 §19.2.1.16 (``p:custShowLst``); sections are
`[MS-PPTX] <https://learn.microsoft.com/openspecs/office_standards/ms-pptx/>`_
§2.1.20, carried in the standard's extension mechanism rather than the
standard itself.

SPEC §4.7. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_plus.core.oxml import part_root, remove, xpath
from pptx_plus.core.reltypes import EXT_URI_SECTION_LST

if TYPE_CHECKING:
    from lxml import etree
    from pptx.presentation import Presentation


def section_lst(prs: Presentation) -> etree._Element | None:
    """Return the deck's ``<p14:sectionLst>``, or None if it has no sections.

    Args:
        prs: The presentation.

    Returns:
        The element, or None.

    Never creates the element. A deck without sections is the common case and
    giving it an empty section list would be an unrequested change to the
    saved bytes (SPEC §9.8).
    """
    found = xpath(
        part_root(prs.part),
        "./p:extLst/p:ext[@uri=$uri]/p14:sectionLst",
        uri=EXT_URI_SECTION_LST,
    )
    return found[0] if found else None


def custom_show_lst(prs: Presentation) -> etree._Element | None:
    """Return the deck's ``<p:custShowLst>``, or None if it has no custom shows.

    Args:
        prs: The presentation.

    Returns:
        The element, or None.
    """
    found = xpath(part_root(prs.part), "./p:custShowLst")
    return found[0] if found else None


def scrub_slide(prs: Presentation, *, slide_id: int, r_id: str) -> None:
    """Remove every reference to a slide from the sections and custom shows.

    Args:
        prs: The presentation.
        slide_id: The slide's ``p:sldId/@id`` — how sections name it.
        r_id: The slide's relationship id — how custom shows name it.

    Both identifiers are required because the two structures key on different
    ones; see the module docstring. Call this *before* dropping the
    relationship, while both are still resolvable.

    An emptied section or custom show is **left in place**. It is a named
    thing the user created, and an empty one is schema-valid; deleting it
    would be a second, unrequested edit. PowerPoint does the same.
    """
    sections = section_lst(prs)
    if sections is not None:
        for entry in xpath(
            sections,
            "./p14:section/p14:sldIdLst/p14:sldId[@id=$id]",
            id=str(slide_id),
        ):
            remove(entry)

    shows = custom_show_lst(prs)
    if shows is not None:
        for entry in xpath(shows, "./p:custShow/p:sldLst/p:sld[@r:id=$rid]", rid=r_id):
            remove(entry)


def reorder_slide(prs: Presentation, *, slide_id: int, to_index: int) -> None:
    """Relocate a slide's section entry to match its new position in the deck.

    Args:
        prs: The presentation.
        slide_id: The slide's ``p:sldId/@id``.
        to_index: The slide's index in the resulting deck.

    Sections partition the running order into contiguous runs, so a slide that
    moves across a section boundary changes section — which is what PowerPoint
    does when a slide is dragged in the slide sorter, and what keeps the
    sections a partition rather than an arbitrary grouping.

    **Exactly one slide's membership changes: the one that moved.** The source
    section shrinks and the destination section grows. The rejected
    alternative holds each section's size fixed, leaving the boundary at a
    fixed *position* — under which moving slide 1 into section 2 would drag
    slide 3 back into section 1. Nobody dragging a slide expects a different
    slide to change section as a result.

    Nothing happens if the deck has no sections, or if this slide has no
    section entry. The latter is an anomaly in decks assembled by other tools;
    inventing an entry would be a change this operation did not declare
    (SPEC §9.9), and it is not this function's business to repair a deck it was
    asked only to reorder.

    Note:
        Landing exactly on a section boundary is genuinely ambiguous — the
        same running order is produced by appending to one section or
        prepending to the next. The tie goes to the section the slide came
        from, which is what makes ``move_slide(prs, i, i)`` a no-op for the
        sections as well as for the slide order. Any other tie-break silently
        re-sections a deck on a call that changed nothing.
    """
    sections_root = section_lst(prs)
    if sections_root is None:
        return

    entries = xpath(
        sections_root,
        "./p14:section/p14:sldIdLst/p14:sldId[@id=$id]",
        id=str(slide_id),
    )
    if not entries:
        return

    entry = entries[0]
    origin = entry.getparent()
    remove(entry)

    id_lists = xpath(sections_root, "./p14:section/p14:sldIdLst")
    offset = to_index
    for position, id_list in enumerate(id_lists):
        count = len(id_list)
        is_last = position == len(id_lists) - 1
        if offset < count or (offset == count and (id_list is origin or is_last)):
            id_list.insert(offset, entry)
            return
        offset -= count

    # Only reachable when the sections do not cover the whole deck, which
    # means the deck arrived that way. Put the entry back where it was rather
    # than inventing a placement.
    origin.append(entry)


__all__ = [
    "custom_show_lst",
    "reorder_slide",
    "scrub_slide",
    "section_lst",
]
