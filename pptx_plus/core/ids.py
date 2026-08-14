"""Slide-id allocation and validation.

A ``<p:sldId>`` carries a deck-scoped ``@id`` that is **not** its relationship
id and **not** its index (SPEC §3.3). ECMA-376's ``ST_SlideId`` constrains it
to ``[256, 2147483647]``, and PowerPoint rejects a deck that violates that.

Allocation delegates to python-pptx's ``CT_SlideIdList.add_sldId``, whose
primary branch is ``max(used) + 1``. That behaviour is worth stating as a
contract rather than treating as an implementation detail:

    **Slide ids are never reused.** A duplicate gets an id distinct from its
    source's, and an id freed by a delete is not handed out again in the same
    session. So anything holding a slide id across a delete gets a clean miss
    rather than silently resolving to a different slide.

SPEC §4, §5.6. This module imports only from ``pptx_plus.core`` (SPEC §9.1).
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx_plus.core.errors import PptxPlusError
from pptx_plus.core.oxml import xpath

if TYPE_CHECKING:
    from lxml import etree

#: Lowest legal ``p:sldId/@id``. ECMA-376 Part 1, ``ST_SlideId``.
MIN_SLIDE_ID = 256
#: Highest legal ``p:sldId/@id``. ECMA-376 Part 1, ``ST_SlideId``.
MAX_SLIDE_ID = 2147483647


class SlideIdRangeError(PptxPlusError, ValueError):
    """Raised for a ``p:sldId/@id`` outside ``[256, 2147483647]``.

    Subclasses ``ValueError`` so existing ``except ValueError:`` clauses still
    catch it; also subclasses :class:`~pptx_plus.core.errors.PptxPlusError`
    per SPEC §9.7.
    """


def validate_slide_id(value: int) -> int:
    """Return ``value`` if it is a legal slide id, else raise.

    Args:
        value: A candidate ``p:sldId/@id``.

    Returns:
        ``value``, unchanged.

    Raises:
        SlideIdRangeError: The value is outside ``[256, 2147483647]``.

    python-pptx enforces the range on *write*, through its ``ST_SlideId``
    attribute converter, but not on read. Decks produced by third-party
    writers do turn up with out-of-range ids, so the integrity battery
    validates what it loads rather than assuming the writer was careful.
    """
    if not MIN_SLIDE_ID <= value <= MAX_SLIDE_ID:
        raise SlideIdRangeError(
            f"slide id {value} is outside the legal range "
            f"[{MIN_SLIDE_ID}, {MAX_SLIDE_ID}] (ECMA-376 ST_SlideId)"
        )
    return value


def used_slide_ids(sld_id_lst: etree._Element) -> list[int]:
    """Return every ``p:sldId/@id`` in ``sld_id_lst``, in document order.

    Args:
        sld_id_lst: A ``<p:sldIdLst>`` element.

    Returns:
        The ids as integers. Duplicates are preserved rather than collapsed —
        a duplicate id is itself a defect the caller may want to report, and
        de-duplicating here would hide it.
    """
    return [int(value) for value in xpath(sld_id_lst, "./p:sldId/@id")]


def next_slide_id(sld_id_lst: etree._Element) -> int:
    """Return the id python-pptx would allocate for the next slide.

    Args:
        sld_id_lst: A ``<p:sldIdLst>`` element.

    Returns:
        ``max(used) + 1``, or :data:`MIN_SLIDE_ID` for an empty list.

    Raises:
        SlideIdRangeError: The deck has exhausted the legal id space.

    This mirrors ``CT_SlideIdList._next_id`` rather than replacing it —
    allocation itself goes through ``add_sldId``. It exists so callers can
    *predict* the next id without mutating anything, which is what lets the
    tests assert the never-reused contract directly.

    Note:
        Upstream's overflow fallback compares an enumeration ordinal against
        an id value and does not do what its name suggests. It is unreachable
        without an existing id at :data:`MAX_SLIDE_ID`, and correcting it is
        upstream's business — but this function raises rather than silently
        returning a colliding id, so the condition surfaces here first.
    """
    used = used_slide_ids(sld_id_lst)
    if not used:
        return MIN_SLIDE_ID
    candidate = max(used) + 1
    if candidate > MAX_SLIDE_ID:
        raise SlideIdRangeError(
            f"cannot allocate a slide id: the deck already uses {MAX_SLIDE_ID}, "
            f"the maximum permitted by ECMA-376 ST_SlideId"
        )
    return candidate


__all__ = [
    "MAX_SLIDE_ID",
    "MIN_SLIDE_ID",
    "SlideIdRangeError",
    "next_slide_id",
    "used_slide_ids",
    "validate_slide_id",
]
