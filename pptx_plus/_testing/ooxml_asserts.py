"""Format invariants for a saved .pptx package.

Every assertion here reads the **saved zip**: its part names, its
``[Content_Types].xml``, its ``_rels/*.rels`` files, and the XML it actually
contains. Nothing in this module touches a live ``Presentation`` beyond
serializing one, and that is the whole point.

Per SPEC §3.5 the in-memory object graph keeps stale references after a delete
by design, while the serialized package is clean. An assertion made against
the model therefore grades the wrong artifact — it can pass on a package that
PowerPoint refuses to open, and fail on one that is perfectly well-formed.
Reading the zip also means these assertions hold for a deck this library never
touched, which is what lets them be graded against known-good and known-broken
input before they are trusted to grade a verb (SPEC §10.5).

What is encoded here is **format invariants**, not test-case expectations. An
assertion belongs in this module only if violating it makes the package wrong
for every caller, not merely different from what one test wanted.

The invariant battery is SPEC §10.3:

===  ==========================================================================
I1   Every ``p:sldId/@r:id`` resolves to a slide relationship
I2   Slide-relationship count equals ``len(sldIdLst)``
I3   Every ``p:sldId/@id`` is unique and in ``[256, 2147483647]``
I4   Every relationship-id attribute in every reachable part resolves
I5   Part names are unique across the package
I6   Section and custom-show entries name live slides
I7   After a duplicate: owned parts disjoint, shared parts identical
===  ==========================================================================

SPEC §10.5.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass
from typing import TYPE_CHECKING

from lxml import etree
from pptx import Presentation as open_presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from pptx_plus.core.ids import MAX_SLIDE_ID, MIN_SLIDE_ID
from pptx_plus.core.ns import qn
from pptx_plus.core.oxml import xpath
from pptx_plus.core.reltypes import (
    EXT_URI_SECTION_LST,
    SCOPE_SELF,
    UNQUALIFIED_REL_ID_ATTRS,
)

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

    from pptx.presentation import Presentation

#: The shape of a relationship id, used by the unclaimed-literal detector.
_RID_LITERAL = re.compile(r"^rId\d+$")

#: The `r:` namespace in Clark-notation prefix form. Every attribute in it is a
#: relationship id -- the namespace is closed by schema, which is what makes a
#: sweep safe (SPEC §4.4).
_R_PREFIX = "{" + qn("r:id")[1:].partition("}")[0] + "}"

_CONTENT_TYPES_PART = "/[Content_Types].xml"


# ---------------------------------------------------------------------------
# Reading a saved package
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Rel:
    """One relationship, as it appears in a `.rels` part."""

    r_id: str
    reltype: str
    target: str
    is_external: bool
    #: The absolute part name the target resolves to, or None when external.
    partname: str | None


class SavedPackage:
    """A saved .pptx, read as a zip rather than through python-pptx.

    Part names are normalized to the OPC form with a leading slash
    (``/ppt/presentation.xml``), which is what `.rels` targets resolve to and
    what ``[Content_Types].xml`` overrides are keyed on. Zip entry names carry
    no leading slash, so the two are not interchangeable.
    """

    def __init__(self, blob: bytes) -> None:
        self.blob = blob
        self._zip = zipfile.ZipFile(io.BytesIO(blob))
        #: Every zip entry, in order, WITH duplicates preserved -- a zip may
        #: legally contain the same name twice and that is exactly what I5 is
        #: looking for, so `namelist()` is used rather than a set.
        self.entries: tuple[str, ...] = tuple(self._zip.namelist())
        self.partnames: tuple[str, ...] = tuple("/" + name for name in self.entries)
        self._xml_cache: dict[str, etree._Element] = {}
        self._rels_cache: dict[str, dict[str, Rel]] = {}
        self._referrer_cache: dict[str, set[str]] | None = None

    # -- raw access --------------------------------------------------------

    def __contains__(self, partname: str) -> bool:
        return partname in self.partnames

    def blob_of(self, partname: str) -> bytes:
        """Return the raw bytes of a part."""
        if partname not in self.partnames:
            raise AssertionError(f"{partname} is not in the package")
        return self._zip.read(partname.lstrip("/"))

    def xml(self, partname: str) -> etree._Element:
        """Return the parsed root element of an XML part, cached."""
        if partname not in self._xml_cache:
            self._xml_cache[partname] = etree.fromstring(self.blob_of(partname))
        return self._xml_cache[partname]

    # -- content types -----------------------------------------------------

    def content_type(self, partname: str) -> str | None:
        """Return the content type declared for a part, or None if undeclared.

        An Override wins over a Default; a Default is keyed on the extension,
        case-insensitively per the OPC spec.
        """
        root = self.xml(_CONTENT_TYPES_PART)
        overrides = xpath(root, "./ct:Override[@PartName=$name]/@ContentType", name=partname)
        if overrides:
            return str(overrides[0])
        extension = _extension_of(partname)
        if not extension:
            return None
        defaults = xpath(root, "./ct:Default/@Extension")
        for declared in defaults:
            if str(declared).lower() == extension:
                found = xpath(
                    root,
                    "./ct:Default[@Extension=$ext]/@ContentType",
                    ext=str(declared),
                )
                return str(found[0])
        return None

    # -- relationships -----------------------------------------------------

    @staticmethod
    def rels_partname_for(partname: str) -> str:
        """Return the name of the `.rels` part describing ``partname``.

        ``"/"`` denotes the package root, whose relationships live in
        ``/_rels/.rels``.
        """
        if partname == "/":
            return "/_rels/.rels"
        directory, _, base = partname.rpartition("/")
        return f"{directory}/_rels/{base}.rels"

    def rels(self, partname: str) -> dict[str, Rel]:
        """Return ``{rId: Rel}`` for a part, empty when it has no `.rels`.

        A part with no relationships legitimately has no `.rels` part at all,
        so absence is not a failure.
        """
        if partname in self._rels_cache:
            return self._rels_cache[partname]

        rels_partname = self.rels_partname_for(partname)
        result: dict[str, Rel] = {}
        if rels_partname in self.partnames:
            base = "/" if partname == "/" else posixpath.dirname(partname)
            for node in xpath(self.xml(rels_partname), "./pr:Relationship"):
                r_id = str(node.get("Id"))
                target = str(node.get("Target"))
                is_external = node.get("TargetMode") == "External"
                result[r_id] = Rel(
                    r_id=r_id,
                    reltype=str(node.get("Type")),
                    target=target,
                    is_external=is_external,
                    partname=None if is_external else _resolve_target(base, target),
                )
        self._rels_cache[partname] = result
        return result

    def targets(self, partname: str, reltypes: Iterable[str] | None = None) -> list[Rel]:
        """Return a part's relationships, optionally filtered by type."""
        wanted = None if reltypes is None else frozenset(reltypes)
        return [
            rel for rel in self.rels(partname).values() if wanted is None or rel.reltype in wanted
        ]

    # -- the part graph ----------------------------------------------------

    def iter_reachable(self) -> Iterator[str]:
        """Yield every part name reachable from the package root.

        The same walk ``OpcPackage.save`` performs, which is what makes
        "reachable" and "written" the same set for a package this library
        produced -- and what makes a *difference* between them detectable in a
        package produced by something else.
        """
        seen: set[str] = set()
        queue = [rel.partname for rel in self.rels("/").values() if rel.partname]
        while queue:
            partname = queue.pop()
            if partname in seen:
                continue
            seen.add(partname)
            yield partname
            queue.extend(rel.partname for rel in self.rels(partname).values() if rel.partname)

    @property
    def presentation_partname(self) -> str:
        """Return the name of the presentation part."""
        for rel in self.rels("/").values():
            if rel.reltype == RT.OFFICE_DOCUMENT and rel.partname:
                return rel.partname
        raise AssertionError("package root has no officeDocument relationship")

    @property
    def presentation(self) -> etree._Element:
        """Return the parsed ``<p:presentation>`` element."""
        return self.xml(self.presentation_partname)

    @property
    def slide_partnames(self) -> list[str]:
        """Return slide part names in presentation order.

        Order comes from ``sldIdLst``, never from the part name -- a part
        called ``slide3.xml`` may be first in the deck, or not in it at all.
        """
        rels = self.rels(self.presentation_partname)
        names = []
        for r_id in xpath(self.presentation, "./p:sldIdLst/p:sldId/@r:id"):
            rel = rels.get(str(r_id))
            if rel is not None and rel.partname:
                names.append(rel.partname)
        return names


def _extension_of(partname: str) -> str:
    """Return a part name's OPC extension, lowercased.

    Not ``posixpath.splitext``. OPC defines the extension as the text after
    the final ``.`` of the final segment, so ``/_rels/.rels`` has extension
    ``rels`` -- while ``splitext`` treats a leading dot as a hidden-file
    marker and reports no extension at all. Every ``.rels`` part in the
    package resolves through the ``rels`` Default, so getting this wrong makes
    the content-type check fail on every valid deck.
    """
    segment = partname.rpartition("/")[2]
    _, dot, extension = segment.rpartition(".")
    return extension.lower() if dot else ""


def _resolve_target(base: str, target: str) -> str:
    """Resolve a relationship target against the containing part's base URI.

    Targets are relative to the *part's* directory, so ``../media/image1.png``
    from ``/ppt/slides/slide1.xml`` is ``/ppt/media/image1.png``. Always
    forward-slashed: these are zip member paths, not filesystem paths, on
    every platform.
    """
    if target.startswith("/"):
        return posixpath.normpath(target)
    return posixpath.normpath(posixpath.join(base, target))


def saved(prs: Presentation) -> SavedPackage:
    """Serialize a presentation and return it as a readable package."""
    stream = io.BytesIO()
    prs.save(stream)
    return SavedPackage(stream.getvalue())


def roundtrip(prs: Presentation) -> Presentation:
    """Save a presentation to memory and reopen it.

    SPEC §10.2: assertions about an operation's effect run against the
    reopened deck, because the in-memory graph and the serialized package do
    not agree after a delete and only the latter is the artifact a user gets.
    """
    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return open_presentation(stream)


# ---------------------------------------------------------------------------
# I1, I2 -- slide list and slide relationships agree
# ---------------------------------------------------------------------------


def assert_slide_rels_consistent(pkg: SavedPackage) -> None:
    """I1 + I2: ``sldIdLst`` and the presentation part's slide rels agree.

    Both directions matter and they fail differently. A ``sldId`` naming a
    missing relationship is a dangling reference PowerPoint reports as damage;
    a slide relationship with no ``sldId`` is the naive-delete signature --
    the slide part stays reachable, so it is still written, but nothing shows
    it. The deck opens with a slide that has silently vanished from the
    running order while still inflating the file.
    """
    presentation_part = pkg.presentation_partname
    rels = pkg.rels(presentation_part)

    referenced = [str(r_id) for r_id in xpath(pkg.presentation, "./p:sldIdLst/p:sldId/@r:id")]
    for r_id in referenced:
        rel = rels.get(r_id)
        if rel is None:
            raise AssertionError(
                f"I1: p:sldId/@r:id={r_id!r} has no relationship on {presentation_part}"
            )
        if rel.reltype != RT.SLIDE:
            raise AssertionError(
                f"I1: p:sldId/@r:id={r_id!r} resolves to a {rel.reltype!r} "
                f"relationship, not a slide"
            )

    slide_rel_ids = {rel.r_id for rel in rels.values() if rel.reltype == RT.SLIDE}
    orphaned = slide_rel_ids - set(referenced)
    if orphaned:
        raise AssertionError(
            f"I2: {len(orphaned)} slide relationship(s) on {presentation_part} have no "
            f"p:sldId: {sorted(orphaned)}. The slide parts are still reachable and were "
            f"written to the package, but no longer appear in the deck."
        )


# ---------------------------------------------------------------------------
# I3 -- slide ids
# ---------------------------------------------------------------------------


def assert_slide_ids_valid(pkg: SavedPackage) -> None:
    """I3: every ``p:sldId/@id`` is unique and within ``ST_SlideId``.

    Uniqueness is what makes a slide id usable as a handle at all, and the
    range is what PowerPoint enforces on open. python-pptx validates the range
    on write but not on read, so a deck from another writer can carry an
    out-of-range id into an operation that then propagates it.
    """
    ids = [int(value) for value in xpath(pkg.presentation, "./p:sldIdLst/p:sldId/@id")]

    duplicates = sorted({value for value in ids if ids.count(value) > 1})
    if duplicates:
        raise AssertionError(f"I3: duplicate p:sldId/@id values: {duplicates}")

    for value in ids:
        if not MIN_SLIDE_ID <= value <= MAX_SLIDE_ID:
            raise AssertionError(
                f"I3: slide id {value} is outside the legal range "
                f"[{MIN_SLIDE_ID}, {MAX_SLIDE_ID}] (ECMA-376 ST_SlideId)"
            )


# ---------------------------------------------------------------------------
# I4 -- every relationship-id attribute resolves
# ---------------------------------------------------------------------------


def _iter_rel_id_attrs(root: etree._Element) -> Iterator[tuple[etree._Element, str, str, str]]:
    """Yield ``(element, attribute name, value, scope)`` for every rel-id attribute.

    Two sources, and the asymmetry between them is deliberate. The ``r:``
    namespace is swept wholesale because it is closed by schema -- nine
    attributes, every one a relationship reference -- so a sweep has no false
    positives by construction. Everything else has to be registered by name,
    because those namespaces are open.

    The scope says which part's relationships the value resolves against:
    its own (the universal rule) or the referring part's.
    """
    for node in root.iter():
        if not isinstance(node.tag, str):
            continue  # comment or processing instruction
        for key, value in node.attrib.items():
            name = str(key)
            if name.startswith(_R_PREFIX):
                yield node, name, str(value), SCOPE_SELF
            else:
                scope = UNQUALIFIED_REL_ID_ATTRS.get((node.tag, name))
                if scope is not None:
                    yield node, name, str(value), scope


def assert_rel_ids_resolve(pkg: SavedPackage) -> None:
    """I4: every relationship-id attribute in every reachable part resolves.

    This is the assertion the whole library exists to satisfy. A dangling
    ``r:embed`` is precisely what ``deepcopy`` of a slide's XML produces
    (SPEC §3.6), and its symptom is a picture that silently fails to render --
    no error, no repair prompt, just a missing image.

    An empty value is not a dangling reference. ``<a:hlinkClick r:id="">`` is
    the ordinary encoding of an action-only hyperlink, and a real deck is full
    of them.

    A **parent-scoped** id resolves against the relationships of whichever
    part refers to the one holding it, not its own. A SmartArt data part has
    no relationships at all and still carries a ``relId``; checking that
    against its own (empty) set would report every PowerPoint-authored diagram
    as damaged. Resolving it in any referring part is enough -- a data part
    referenced by two slides is legal, and the id is meaningful in each.
    """
    referrers = _referrers(pkg)
    for partname in pkg.iter_reachable():
        content_type = pkg.content_type(partname)
        if not _is_xml(content_type):
            continue
        # Both id sets are computed once per part, not per attribute: a busy
        # slide has hundreds of `r:` attributes and rebuilding the set for
        # each one dominated the whole test suite's runtime.
        own = frozenset(pkg.rels(partname))
        inherited: frozenset[str] | None = None

        for node, name, value, scope in _iter_rel_id_attrs(pkg.xml(partname)):
            if not value:
                continue
            if scope == SCOPE_SELF:
                if value in own:
                    continue
                where = f"{partname} declares only {sorted(own)}"
            else:
                if inherited is None:
                    sources: set[str] = referrers.get(partname, set())
                    inherited = frozenset(r_id for source in sources for r_id in pkg.rels(source))
                if value in inherited:
                    continue
                where = (
                    f"no part referring to it "
                    f"({sorted(referrers.get(partname, set()))}) declares it"
                )
            raise AssertionError(
                f"I4: {partname} has a dangling relationship reference: "
                f"<{etree.QName(node).localname} {name}={value!r}> but {where}"
            )


def _referrers(pkg: SavedPackage) -> dict[str, set[str]]:
    """Return ``{partname: parts that hold a relationship to it}``, memoized.

    Needed only to resolve parent-scoped ids, but the index is over the whole
    graph, and the battery calls this assertion on every package it grades --
    including once per generated example in the property tests. Memoized on
    the package because a `SavedPackage` is immutable: it wraps one fixed
    blob, so the answer cannot change.
    """
    if pkg._referrer_cache is None:
        result: dict[str, set[str]] = {}
        for source in ("/", *pkg.iter_reachable()):
            for rel in pkg.rels(source).values():
                if rel.partname:
                    result.setdefault(rel.partname, set()).add(source)
        pkg._referrer_cache = result
    return pkg._referrer_cache


def assert_no_unclaimed_rid_literals(pkg: SavedPackage) -> None:
    """The tripwire for a relationship-id attribute nobody knows about.

    Any attribute whose value is shaped like ``rId7`` but which is neither in
    the ``r:`` namespace nor in the registry is reported. It is the companion
    to :func:`assert_rel_ids_resolve`: that function proves the ids we *know*
    about resolve, and this one looks for ids we do not know about at all.

    Together they are the argument for a namespace sweep over an element
    allowlist, turned executable. An allowlist keyed on element name is keyed
    on an open vocabulary that grows with every Office release and fails
    silently on each addition; this pair fails loudly instead, which is how
    the next ``dsp:dataModelExt/@relId`` gets found.

    A false positive is possible in principle -- a user could name a shape
    ``rId7`` -- which is why at runtime the equivalent check is a warning and
    only the test suite treats it as fatal (SPEC §4.4).
    """
    unclaimed: list[str] = []
    for partname in pkg.iter_reachable():
        if not _is_xml(pkg.content_type(partname)):
            continue
        for node in pkg.xml(partname).iter():
            if not isinstance(node.tag, str):
                continue
            for key, value in node.attrib.items():
                name = str(key)
                if name.startswith(_R_PREFIX) or (node.tag, name) in UNQUALIFIED_REL_ID_ATTRS:
                    continue
                if _RID_LITERAL.match(str(value)):
                    unclaimed.append(
                        f"{partname}: <{etree.QName(node).localname} {name}={value!r}>"
                    )
    if unclaimed:
        raise AssertionError(
            "unclaimed relationship-id literal(s) -- an attribute holding what looks "
            "like a relationship id but which the rewriter does not know to rewrite. "
            "Either add it to UNQUALIFIED_REL_ID_ATTRS or confirm it is a coincidence:\n  "
            + "\n  ".join(unclaimed)
        )


def _is_xml(content_type: str | None) -> bool:
    """Return whether a content type denotes XML this harness should parse."""
    if content_type is None:
        return False
    return content_type.endswith("+xml") or content_type == "application/xml"


# ---------------------------------------------------------------------------
# I5 -- part names
# ---------------------------------------------------------------------------


def assert_partnames_unique(pkg: SavedPackage) -> None:
    """I5: no part name appears twice in the package.

    A zip file can hold two entries with the same name; most readers return
    the last one and the earlier is simply lost. This is the failure mode of
    allocating a part name without a reservation set while cloning a slide
    that owns two charts (SPEC §4.6) -- both clones are handed
    ``/ppt/charts/chart2.xml`` and one silently disappears.
    """
    duplicates = sorted({name for name in pkg.entries if pkg.entries.count(name) > 1})
    if duplicates:
        raise AssertionError(f"I5: duplicate part name(s) in the package: {duplicates}")


def assert_content_types_complete(pkg: SavedPackage) -> None:
    """Every part in the package has a declared content type.

    A part with no Override and no matching Default extension is unopenable:
    the consumer cannot know how to interpret it, and PowerPoint reports the
    package as damaged rather than ignoring the part.

    The limit is worth stating. An OPC package declares
    ``Default Extension="xml"``, so *any* ``.xml`` part resolves to at least
    ``application/xml`` and this assertion cannot detect a slide that lost its
    specific Override. It catches the case that actually occurs -- a part
    written under an extension nobody declared. Detecting a wrong-but-present
    content type would need a reltype-to-content-type table, which
    ``[Content_Types].xml`` being regenerated from the part tuple at save
    (SPEC §3.5) makes unreachable from this library's own output.
    """
    missing = [
        partname
        for partname in pkg.partnames
        if partname != _CONTENT_TYPES_PART and pkg.content_type(partname) is None
    ]
    if missing:
        raise AssertionError(f"no content type declared for: {sorted(missing)}")


def assert_all_xml_parses(pkg: SavedPackage) -> None:
    """Every part declared as XML actually parses.

    Cheap, runs everywhere, and catches most of what a headless LibreOffice
    conversion would -- at none of its setup cost.
    """
    for partname in pkg.partnames:
        if partname == _CONTENT_TYPES_PART or _is_xml(pkg.content_type(partname)):
            try:
                pkg.xml(partname)
            except etree.XMLSyntaxError as exc:
                raise AssertionError(f"{partname} is not well-formed XML: {exc}") from exc


# ---------------------------------------------------------------------------
# I6 -- sections and custom shows
# ---------------------------------------------------------------------------


def assert_sections_consistent(pkg: SavedPackage) -> None:
    """I6: every section and custom-show entry names a live slide.

    The highest-severity, lowest-visibility invariant in the battery. Neither
    structure is modelled by python-pptx, and neither appears in any deck a
    test can generate -- they survive round trips only because unrecognized
    XML is preserved verbatim. So an implementation can pass every other
    assertion here and still produce a repair prompt on the first deck a real
    user has organized into sections.

    Sections key on ``sldId/@id`` (the deck-scoped slide id); custom shows key
    on ``@r:id`` (a relationship id on the presentation part). Two different
    identifiers for the same slide, in two structures a hundred lines apart in
    the same file -- see SPEC §3.3.
    """
    presentation = pkg.presentation
    live_ids = {int(value) for value in xpath(presentation, "./p:sldIdLst/p:sldId/@id")}
    live_rids = {str(value) for value in xpath(presentation, "./p:sldIdLst/p:sldId/@r:id")}

    section_ids = xpath(
        presentation,
        "./p:extLst/p:ext[@uri=$uri]/p14:sectionLst/p14:section/p14:sldIdLst/p14:sldId/@id",
        uri=EXT_URI_SECTION_LST,
    )
    stale_sections = sorted({int(value) for value in section_ids} - live_ids)
    if stale_sections:
        raise AssertionError(
            f"I6: p14:sectionLst references slide id(s) not in the deck: {stale_sections}. "
            f"PowerPoint reports this as damage and offers to repair the file."
        )

    show_rids = xpath(presentation, "./p:custShowLst/p:custShow/p:sldLst/p:sld/@r:id")
    stale_shows = sorted({str(value) for value in show_rids} - live_rids)
    if stale_shows:
        raise AssertionError(
            f"I6: p:custShowLst references relationship id(s) that are no longer "
            f"slides in the deck: {stale_shows}"
        )


# ---------------------------------------------------------------------------
# I7 -- what a duplicate shares and what it owns
# ---------------------------------------------------------------------------


def related_partnames(
    pkg: SavedPackage,
    partname: str,
    reltypes: Iterable[str] | None = None,
) -> set[str]:
    """Return the part names a part relates to, optionally filtered by type."""
    return {rel.partname for rel in pkg.targets(partname, reltypes) if rel.partname}


def assert_parts_shared(
    pkg: SavedPackage,
    first: str,
    second: str,
    *,
    reltypes: Iterable[str],
) -> None:
    """I7: two parts reference the *same* targets of the given types.

    Sharing is the format's own encoding of "the same picture": an image part
    is identified by its bytes, so cloning it would inflate the package to say
    something the format already says better.
    """
    left = related_partnames(pkg, first, reltypes)
    right = related_partnames(pkg, second, reltypes)
    if not left:
        raise AssertionError(f"I7: {first} has no relationships of the given types to share")
    if left != right:
        raise AssertionError(
            f"I7: expected {first} and {second} to share targets, but\n"
            f"  {first} -> {sorted(left)}\n"
            f"  {second} -> {sorted(right)}"
        )


def assert_parts_disjoint(
    pkg: SavedPackage,
    first: str,
    second: str,
    *,
    reltypes: Iterable[str],
) -> None:
    """I7: two parts own *separate* targets of the given types.

    The converse of sharing, and the one that actually needs proving. A chart,
    a notes slide, and a diagram's definition parts are mutable units: if a
    duplicate shares them, editing the copy silently edits the original, which
    is the kind of bug that surfaces weeks later as "PowerPoint changed my
    other slide."
    """
    left = related_partnames(pkg, first, reltypes)
    right = related_partnames(pkg, second, reltypes)
    if not left or not right:
        raise AssertionError(
            f"I7: expected both parts to own targets of the given types, but\n"
            f"  {first} -> {sorted(left)}\n"
            f"  {second} -> {sorted(right)}"
        )
    overlap = left & right
    if overlap:
        raise AssertionError(
            f"I7: {first} and {second} share owned part(s) that must be distinct: {sorted(overlap)}"
        )


# ---------------------------------------------------------------------------
# Presence
# ---------------------------------------------------------------------------


def assert_in_package(pkg: SavedPackage, partname: str) -> None:
    """The named part was written to the package."""
    if partname not in pkg:
        raise AssertionError(f"{partname} is not in the saved package")


def assert_not_in_package(pkg: SavedPackage, partname: str) -> None:
    """The named part was not written to the package.

    Checked on the saved zip rather than the in-memory graph, because after a
    delete the in-memory graph still holds the part by design (SPEC §3.5) --
    asserting there would fail on a perfectly correct delete.
    """
    if partname in pkg:
        raise AssertionError(f"{partname} is still in the saved package")


# ---------------------------------------------------------------------------
# The battery
# ---------------------------------------------------------------------------


def assert_package_integrity(pkg: SavedPackage) -> None:
    """Run every invariant that holds for any valid package: I1-I6.

    I7 is excluded because it is a claim about the *relationship between* two
    parts after a specific operation, not a property a package has on its own.

    Ordered cheapest and most fundamental first, so the first failure reported
    is the most explanatory one: a package whose XML does not parse will fail
    several later assertions for reasons that all trace back to that.
    """
    assert_partnames_unique(pkg)
    assert_content_types_complete(pkg)
    assert_all_xml_parses(pkg)
    assert_slide_rels_consistent(pkg)
    assert_slide_ids_valid(pkg)
    assert_rel_ids_resolve(pkg)
    assert_no_unclaimed_rid_literals(pkg)
    assert_sections_consistent(pkg)


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
