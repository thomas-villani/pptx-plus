# pptx_plus — API Specification

This document is the **contract as it stands today**. Implementations that
diverge from the public API specified here are wrong, even if they work.
Internal implementation is flexible; §9 lists the invariants that constrain it.

Three documents, three jobs, and they do not overlap:

| Document | Answers |
|---|---|
| `SPEC.md` (this file) | What the library guarantees **right now** |
| `ROADMAP.md` | What is coming, and what is refused |
| `CHANGELOG.md` | What changed, and when |

Code comments cite `SPEC §N`. Nothing in the code cites ROADMAP — a roadmap
item has no contract until it lands here.

Every section carries a **Status** line. Sections §6 and §7 are reserved and
empty on purpose, so that landing cross-deck copy and SmartArt reading in a
later cycle renumbers nothing.

§9, §10, §13 and §16 keep their numbering aligned with
[`docx-plus`](https://github.com/thomas-villani/docx-plus)'s `SPEC.md`, so
"SPEC §9.1" means the import invariant in both repositories.

---

## 1. Purpose & Non-Goals

**Status:** stable.

### 1.1 What pptx_plus is

The library every python-pptx power user ends up writing badly: hardened
helpers for the OPC operations that sit just past python-pptx's abstraction
boundary.

python-pptx models a presentation's *content* well and its *package* barely at
all. The gap that matters is the slide lifecycle — deleting, reordering, and
duplicating slides — because doing it correctly is not a matter of adding a
convenience method. It requires rewriting relationships and parts, and every
recipe in circulation skips exactly that step.

v0.1 targets one capability:

- **Slide lifecycle** — `delete_slide`, `move_slide`, `duplicate_slide`,
  operating at the OPC layer, with the relationship graph left closed and
  every side-index (sections, custom shows) maintained.

### 1.2 Non-goals for v0.1

Explicitly out of scope, and not to be implemented even where they look like
natural extensions:

- **Animations and transitions.** See §1.3.
- **SmartArt beyond preservation.** Reading node text, editing it, and
  template-clone authoring are roadmap items, not v0.1. What v0.1 provides is
  that a duplicated slide's SmartArt *survives* — a consequence of cloning at
  the part level, not a feature with an API.
- **Chart editing beyond what survives a part-level clone.**
- **Cross-deck copy and merge.** Reserved as §6; targeted for v0.2.
- **A CLI.** The three v0.1 verbs mutate a deck in place and would need
  `--output` / backup semantics before they are safe to expose to a shell.

### 1.3 Stylistic non-goals

Permanent. These are not "not yet."

- **Templating.** Use a templating library over the content API.
- **A `Presentation` subclass**, or any wrapper that replaces python-pptx's
  object model. Callers keep their `Presentation` and reach for `pptx_plus`
  only where python-pptx stops.
- **A SmartArt layout engine.** A SmartArt diagram references a *layout
  definition part* that is not a static template but a declarative layout
  program — algorithm nodes, constraint lists, `forEach`, `choose`/`if`.
  PowerPoint executes it against the data model to compute the drawing cache
  every viewer actually renders. No open implementation of that engine exists.
  Self-computing diagram geometry is refused permanently, and any capability
  that would require it is refused with it.
- **Feature parity with the commercial libraries.** The bar is "useful and
  maintainable by one person."

### 1.4 Relationship to python-pptx, and to the `-ng` fork

pptx_plus **composes with upstream python-pptx 1.0.x**. It is not a fork.

It is specifically *not* built on `python-pptx-ng`. That fork's pitch is that
upstream is unmaintained, which is stale: upstream reached 1.0 and shipped
patches into 2025, while the fork's latest release now predates upstream's
1.0.0. The fork also solves none of the gaps that motivate this project — it
added charts, gradient fills, and table cell merge, and explicitly did not add
slide management or SmartArt. It occupies the same `pptx` import namespace, so
it is an either/or swap rather than a layer. It is MIT, so its chart code can
be cherry-picked later if a need appears.

**A hard fork is deliberately deferred**, not rejected. Revisit only on
concrete evidence of a wall that requires fixing python-pptx's model
internals. Because slide-lifecycle operations run at the OPC layer, below the
object model, python-pptx's lossy re-serialization is never in the loop and
there is nothing to fight — so the usual reason to fork does not apply here.
Architect so this stays a late, evidence-driven choice rather than an upfront
tax.

---

## 2. Architecture Overview

**Status:** stable.

### 2.1 Directory tree

```
pptx_plus/
├── __init__.py          # PptxPlusError, __version__
├── py.typed
├── core/                # foundation — the only legal import target for capabilities
│   ├── errors.py        #   PptxPlusError, zero imports (§9.7)
│   ├── ns.py            #   namespace URIs, NSMAP, BUILD_NSMAP, qn()
│   ├── oxml.py          #   el, sub, xpath, remove, ordered_insert (§9.2)
│   ├── reltypes.py      #   RT/CT constants python-pptx lacks; the clone policy sets
│   ├── ids.py           #   slide-id allocation and range validation
│   ├── parts.py         #   partname allocation, clone_part, drop_relationship
│   ├── relmap.py        #   RelMap, remap_rel_ids (§9.3)
│   ├── partgraph.py     #   read-only part-graph enumeration
│   ├── clone.py         #   ClonePolicy, CloneResult, clone_part_graph
│   ├── sections.py      #   p14:sectionLst and p:custShowLst maintenance
│   └── _compat.py       #   upstream-surface assertion (§14.2)
├── slides/              # the public verbs
│   ├── resolve.py       #   resolve_slide, slide_index, contains
│   ├── delete.py  move.py  duplicate.py
└── _testing/            # test-only assertions; shipped, coverage-omitted
    └── ooxml_asserts.py
```

### 2.2 Dependency invariant

One-way. `core/` depends on nothing above it. Capability packages depend on
`core/` and on nothing else in this library. Enforced by
`tests/test_import_invariant.py`; see §9.1.

### 2.3 Module responsibilities

`core/` is split by **lifecycle stage**, not by data type: enumerate
(`partgraph`) → allocate (`parts`, `ids`) → rewrite (`relmap`) → orchestrate
(`clone`) → maintain deck-level side-indexes (`sections`). Only `clone.py`
mutates the package.

---

## 3. The OPC Model pptx_plus Operates On

**Status:** stable. This section is the vocabulary the rest of the document
uses.

### 3.1 Package, parts, partnames, content types

A `.pptx` is an OPC package: a ZIP of **parts**, each identified by a
**partname** (`/ppt/slides/slide1.xml`) and typed by a **content type**
declared in `[Content_Types].xml`. Partnames are always forward-slashed, on
every platform.

### 3.2 Relationships: `r:id` is part-scoped, not global

Parts reference each other through **relationships**, stored in a sibling
`_rels/<part>.rels` file. A relationship has an `Id` (`rId1`), a `Type`, and a
`Target`.

**The `r:id` referenced inside a part's XML is scoped to that part.** `rId2`
in `slide1.xml` and `rId2` in `slide2.xml` are unrelated. This is the fact
every broken slide-copy recipe gets wrong, and §3.6 spells out how.

### 3.3 The three identifiers called "id"

| Name | Scope | Where | Range |
|---|---|---|---|
| `p:sldId/@id` | deck | `presentation.xml` | 256 … 2147483647 |
| `p:sldId/@r:id` | presentation **part** | same element | `rIdN` |
| list index | call site | `prs.slides[i]` | 0 … len−1 |

They are never interchangeable. Conflating the first two is the second-most
common bug in this problem space.

### 3.4 `presentation.xml`

- `<p:sldIdLst>` — the ordered list of slides. **This is the only thing that
  determines slide order.** A part named `slide3.xml` is not necessarily third
  and may not be in the deck at all.
- `<p:sldLayoutIdLst>` / `<p:sldMasterIdLst>` — layouts and masters are
  referenced from *here*, not from slides, so they survive the deletion of the
  last slide that uses them.
- `<p:custShowLst>` — custom shows, referencing slides by **`r:id`**.
- `<p:extLst>` → `p14:sectionLst` — PowerPoint sections, referencing slides by
  **`sldId/@id`**. python-pptx does not model this at all; it survives round
  trips only because unrecognized XML is preserved verbatim.

### 3.5 Reachability, and what "orphan" means precisely

`OpcPackage.save` serializes `tuple(self.iter_parts())`, and `iter_parts` is a
relationship-graph walk from the package root. `[Content_Types].xml` is
regenerated from that same tuple.

**Therefore a part that nothing references is never written.** Dropping the
presentation part's relationship to a slide is the entire act of collecting
it; the slide part, and any image reachable only through it, simply do not
appear in the saved package.

This is why there is no orphan collector in `core/`. A hand-rolled
mark-and-sweep would be a second, weaker implementation of something the
writer already does correctly. The distinction that matters:

- The **in-memory** object graph keeps stale references after a delete. This
  is why every assertion in the test suite runs against a **saved and
  reopened** package (§10.2).
- The **serialized** package is clean.

### 3.6 Why the common recipes are wrong

**Duplicate.** `deepcopy(slide.element)` + append to `sldIdLst` copies the
slide's `<p:spTree>` and leaves every `r:id` and `r:embed` inside it pointing
at the *source slide's* relationships. Because those ids are part-scoped
(§3.2), they now resolve against the new part's relationship set — where they
mean something else, or nothing. Images, charts, and media are
relationship-targeted parts, so the copy carries dangling references and the
picture silently drops on open.

**Delete.** `prs.slides._sldIdLst.remove(...)` unlinks the `<p:sldId>` but
leaves the presentation part's relationship in place, so the slide part stays
reachable and is still written. It also leaves the slide's entries in
`p14:sectionLst` and `p:custShowLst` — dangling references that produce a
PowerPoint repair prompt.

So the gap is not "no convenience method." **Relationship and part rewriting
is the actual work, and every recipe skips it.**

---

## 4. Core Foundation API — `core/`

**Status:** stable for v0.1. All of `core/` is implemented.

### 4.1 `core/ns.py`

Owns the namespace map. python-pptx's own `pptx.oxml.ns._nsmap` is
insufficient — it lacks `mc` (it binds that URI to `ve`), `dgm`, `dsp`, `p14`,
`a14`, `a16`, and `asvg`, all of which the copy engine must be able to query.

- `NSMAP` — the **query** map. Every prefix the library can address.
- `BUILD_NSMAP` — the **write** map, declared on constructed elements.
  Deliberately tiny (`p`, `a`, `r`): v0.1 is a *rewriter*, not a builder. It
  clones existing XML and edits attribute values in place; the only element it
  ever constructs is `p:sldId`, and python-pptx constructs that. Growth in
  `BUILD_NSMAP` is a signal that a capability module has started authoring,
  and should be questioned.
- `qn(name: str) -> str` — `"p:sldId"` → Clark notation. Memoized. Raises
  `InvalidNamespaceError` on an unknown prefix or a missing colon.

### 4.2 `core/oxml.py`

The single chokepoint for OOXML I/O (§9.2). `el`, `sub`, `xpath`, `remove`,
`ordered_insert`, plus the two private-surface accessors `part_root(part)` and
`sld_id_lst(prs)` that quarantine every use of `_element` (§14.2).

`xpath()` compiles through a cached `etree.XPath` bound to `NSMAP` and takes
**XPath variables**, never f-string interpolation:

```python
xpath(root, "./p:sldId[@r:id=$rid]", rid=r_id)
```

### 4.3 `core/reltypes.py`

Relationship- and content-type constants python-pptx's enums omit — notably
the Microsoft-extension `diagramDrawing` reltype that carries SmartArt's fifth
part — plus the two policy sets that drive the clone engine:

- `SHARE_RELTYPES` — image, media, video, audio, font, thumbnail.
- `STRUCTURAL_RELTYPES` — slideLayout, slideMaster, notesMaster, theme,
  handoutMaster. This is the v0.2 seam (§6).
- `REUSE_RELTYPES` — `RT.SLIDE`. A relationship naming another slide in the
  same deck, as a slide-jump hyperlink does. Never cloned: duplicating a slide
  that links to slide 4 must produce a slide that links to slide 4. Kept
  separate from `STRUCTURAL_RELTYPES` even though both mean "reuse the target"
  today, because structural relationships are the cross-deck seam where the
  layout chain gets *imported* — and a slide-jump target has entirely
  different cross-deck semantics. Folding them together would silently give it
  the wrong ones the moment §6 lands.

  This does **not** cover a notes slide's relationship back to its own slide,
  which is a cycle rather than a reference to a third party. The clone map
  resolves that (§4.5).
- `UNQUALIFIED_REL_ID_ATTRS` — `(element, attribute)` pairs carrying a
  relationship id *outside* the `r:` namespace; see §4.4 for what it is for.
  It lives here rather than in `relmap.py` because it is a constant with two
  consumers: the rewriter that must remap those attributes, and the integrity
  harness that must recognize them as relationship ids rather than report them
  as unclaimed literals. The harness is built first (§10.5), so a registry
  owned by the rewriter would not exist yet when it is first needed.

### 4.4 `core/relmap.py`

`RelMap` maps a source part's `rId`s to the corresponding ids on its clone.

**It is a function, not a bijection.** `_Relationships.get_or_add` dedupes on
`(reltype, target, is_external)`, so two source relationships with the same
type and target collapse to one on the clone and two keys map to one value.
Code that asserts bijectivity is wrong.

`remap_rel_ids(root, relmap, *, strict=True) -> int` rewrites, in place, every
relationship-id-bearing attribute beneath `root`.

**It sweeps the `r:` namespace rather than allowlisting elements.** The `r:`
namespace is closed by schema — `shared-relationshipReference.xsd` declares
exactly nine attributes (`r:id`, `r:embed`, `r:link`, `r:dm`, `r:lo`, `r:qs`,
`r:cs`, `r:href`, `r:pict`) and every relationship-reference attribute group in
OOXML references one of them — so the sweep has zero false positives *by
construction*. An element-keyed allowlist is keyed on an open vocabulary that
grows with every Office release and fails **silently** on the additions.

Two supplements:

- `UNQUALIFIED_REL_ID_ATTRS` (defined in `core/reltypes.py`, §4.3) — a
  registry mapping `(element, attribute)` pairs carrying a relationship id
  *outside* the `r:` namespace to the **scope** their value resolves in. Real
  decks contain exactly one: `dsp:dataModelExt/@relId`, inside a SmartArt data
  part, naming the rendered-drawing part.

  **Its scope is the referring part, not its own**, and that is the half that
  surprises. Every other relationship reference in OOXML is part-scoped
  (§3.2); this one cannot be, because a diagram data part has no relationships
  at all. Its `relId` names one on the *slide* that references the diagram.
  Code that applies the universal rule produces a copy whose diagram points at
  the original's drawing cache — silently, since PowerPoint recomputes the
  drawing on open and looks correct while other renderers do not. Verified
  against a PowerPoint-authored sample, not assumed.
- `RelIdLiteralWarning` — emitted for any attribute value matching `^rId\d+$`
  that was neither rewritten nor registered. A **warning** at runtime, because
  a shape genuinely named `rId7` must not crash a library call; a **hard
  error** in the test suite via `filterwarnings`. This is how the next
  unqualified relationship attribute gets found instead of silently producing a
  dangling reference.

Two mechanical rules, both load-bearing:

1. **Iterate attribute nodes; never string-substitute the serialized blob.**
   A map containing both `rId1 → rId2` and `rId2 → rId1` is a legal outcome of
   a collapse plus a gap-fill, and any textual pass double-applies one of them.
2. **Run the pass per cloned XML part**, not once on the slide. A chart's
   `c:externalData/@r:id` and a diagram's `@relId` live in sub-parts, each with
   its own `RelMap`.

An `r:`-namespace attribute whose value is absent from the map raises
`DanglingRelationshipError` under `strict=True`. Real decks do contain these,
and propagating one silently is worse than failing.

### 4.5 `core/clone.py` — `clone_part_graph`

```python
def clone_part_graph(root: Part, *, into: Package,
                     policy: ClonePolicy = ClonePolicy()) -> CloneResult
```

`into` is a parameter from day one even though v0.1 always passes
`root.package`, so §6 touches two private functions and nothing else.

**The share-vs-deep rule**, stated as a rule rather than a list:

> A related part is **shared by reference** if and only if (a) its
> relationship type is in `SHARE_RELTYPES` — its identity is its bytes, and
> two slides pointing at it is the format's normal encoding of "the same
> picture" — **and** (b) it has zero relationships of its own. Everything else
> reachable and non-structural is **deep-cloned**.

Clause (b) does the real work: anything owning a sub-graph is a mutable unit
and cannot be shared, whatever the type table says.

Charts require no special-casing under this rule. A chart's embedded workbook
is a blob part with no relationships whose type is `package`, not in
`SHARE_RELTYPES` — so it is deep-cloned byte-identically, its relationship
re-minted, and `c:externalData/@r:id` rewritten because it is an `r:`
attribute.

**Three invariants:**

- **Clone-map invariant.** Within one operation each source part maps to at
  most one destination part, and any edge whose target is already mapped
  resolves to the mapping. This makes the notes slide's back-reference to its
  slide correct with no special case, and handles diamonds and cycles.
- The map entry is written **before** recursing. Slide → notesSlide → slide is
  a real cycle in any deck with speaker notes.
- `remap_rel_ids` runs after all of a part's relationships are minted.

### 4.6 `core/parts.py` — partname allocation

`clone_part(src, *, into, reserved)` clones a part byte-faithfully:

```python
type(src).load(new_partname, src.content_type, into, src.blob)
```

`Part.load` and `Part.blob` are public and no python-pptx part class overrides
`load`, so this one expression works uniformly for slides, charts, images,
embedded workbooks, and blob-only parts. For a blob-backed part it copies the
byte string verbatim — no reparse, and therefore no whitespace normalization —
which is what makes the preservation guarantee in §8.1 byte-exact.

**`reserved` is not optional.** `Package.next_partname` derives its used-name
set from `iter_parts()`, a walk of the relationship graph — so a part that has
been constructed but not yet related to anything is invisible to it. Cloning a
slide bearing two charts would allocate the same partname twice and one would
silently overwrite the other. Attaching before recursing is impossible, since
the parent's relationship does not exist yet; the reservation set is the fix.

### 4.7 `core/sections.py` — sections and custom shows

```python
def section_lst(prs) -> _Element | None
def custom_show_lst(prs) -> _Element | None
def scrub_slide(prs, *, slide_id: int, r_id: str) -> None
def reorder_slide(prs, *, slide_id: int, to_index: int) -> None
```

A deck indexes its slides up to three times. `p:sldIdLst` is the running order
and the only one python-pptx models. `p14:sectionLst`, in
`p:presentation/p:extLst`, groups slides into the named sections the slide
sorter shows. `p:custShowLst` names subsets as alternative running orders.

Neither of the latter two is modelled upstream; they survive a round trip only
because unrecognized XML is preserved verbatim. That is what makes them
dangerous — an operation that edits `sldIdLst` and stops leaves them pointing
at a slide that is gone, and PowerPoint reports the file as damaged.

**They key on different identifiers**, which is the trap: sections reference
`p:sldId/@id`, custom shows reference `@r:id`. Same slide, two names, in one
file. §3.3. Hence both parameters on `scrub_slide`; call it while both are
still resolvable, i.e. *before* dropping the relationship.

Two behaviours are contract, not implementation detail:

- **An emptied section or custom show is left in place.** It is a named thing
  the user created and an empty one is schema-valid. PowerPoint does the same.
- **A reorder changes exactly one slide's section membership: the one that
  moved.** Landing on a section boundary is genuinely ambiguous — appending to
  one section and prepending to the next give the same running order — and the
  tie goes to the section the slide came from. That is what makes
  `move_slide(prs, i, i)` a no-op for the sections too. The rejected
  alternative holds section *sizes* fixed, under which moving slide 1 into
  section 2 drags slide 3 back into section 1.

A deck whose sections do not cover every slide is left as found. Repairing it
is not the business of a call that was asked only to reorder (§9.9).

---

## 5. Slide Lifecycle API — `slides/`

**Status:** stable for v0.1. All of §5.1–§5.8 is implemented.

### 5.1 Argument normalization

```python
def resolve_slide(prs, slide_or_index: Slide | int) -> tuple[int, Slide, CT_SlideId]
def slide_index(prs, slide_or_index: Slide | int) -> int
def contains(prs, slide: Slide) -> bool
```

All three verbs need the index, the `Slide`, and the `<p:sldId>` element;
resolving them separately means three walks of `sldIdLst` and three places to
get index arithmetic wrong.

- An `int` follows Python list semantics, negatives included. Out of range
  raises `SlideIndexError`.
- A `Slide` is located by element identity. Not in this deck raises
  `SlideNotFoundError`.
- Anything else raises `TypeError` naming the type received.

### 5.2 `delete_slide`

```python
def delete_slide(prs, slide_or_index: Slide | int) -> None
```

Removes the `<p:sldId>`, drops the presentation part's relationship, and
scrubs the slide from every custom show and every section. The slide part and
everything reachable only through it stop being written on the next save
(§3.5).

Deleting the last slide is **allowed**: an empty `<p:sldIdLst/>` is
schema-valid and PowerPoint opens a zero-slide deck.

### 5.3 `move_slide`

```python
def move_slide(prs, slide_or_index: Slide | int, to_index: int) -> None
```

Pure ordering. No part or relationship changes; the slide's entry relocates
within `p14:sectionLst` if the deck has sections.

### 5.4 `duplicate_slide`

```python
def duplicate_slide(prs, slide_or_index: Slide | int, *,
                    to_index: int | None = None,
                    with_notes: bool = True) -> Slide
```

Clones within the same deck and returns the new `Slide`. Images and media are
shared by reference; charts with their embedded workbooks, SmartArt definition
parts, embedded objects, and — unless `with_notes=False` — the notes slide are
deep-cloned. Every relationship id inside the cloned XML is rewritten.

**`to_index` defaults to immediately after the source**, matching PowerPoint's
own Duplicate Slide and what the word "duplicate" leads a caller to expect.
`to_index=-1` appends. The range is the deck *with* the copy in it, so for a
four-slide deck `0..4` are valid.

A slide-jump hyperlink points at the **same** slide it did before: the copy
links to slide 4, it does not bring a second copy of slide 4 with it. See
`REUSE_RELTYPES` in §4.3.

`with_notes` defaults to `True` because PowerPoint's own Duplicate Slide keeps
speaker notes, and because the costs are asymmetric: losing notes is
unrecoverable, gaining them is one `del` away.

### 5.5 Index semantics

**`to_index` is the position the slide will occupy in the *resulting* deck** —
`list.insert` semantics against the list with the slide already removed.

| Deck | Call | Result |
|---|---|---|
| A B C D | `move_slide(prs, 0, 2)` | B C **A** D |
| A B C D | `move_slide(prs, 2, 2)` | A B C D (no-op) |
| A B C D | `move_slide(prs, 3, 0)` | **D** A B C |

The rejected alternative — "index in the original list" — makes
`move_slide(prs, i, i + 1)` a no-op, which surprises everyone.
`move_slide(prs, i, i)` is a no-op for every valid `i`, and that is the
acceptance criterion that pins the semantics down.

Negative `to_index` counts from the end of the resulting deck.

**Out of range raises `SlideIndexError` rather than clamping.** This diverges
from `list.insert` deliberately: clamping hides off-by-one bugs in caller
loops, which is the bug class this library exists to eliminate.

### 5.6 Idempotency contract

Stated per verb, because the right word differs:

- **`duplicate_slide` is repeatable, not idempotent.** Twice yields two
  distinct slides. What must be idempotent is *allocation*: partname,
  relationship-id, and slide-id allocation are pure functions of package state
  plus the per-operation reservation set, so a second call cannot collide with
  the first.
- **`delete_slide` raises on a second call** rather than silently no-opping.
  `SlideNotFoundError` subclasses `KeyError`, so
  `contextlib.suppress(KeyError)` gives callers opt-in idempotence; a silent
  no-op would hide caller bugs. Note that `delete_slide(prs, 0)` twice deletes
  *two* slides — that is index semantics, not a bug.
- **`move_slide(prs, i, i)` is a no-op** for every valid `i`.

**Slide ids are never reused.** A duplicate gets a fresh `sldId/@id`, and an id
freed by a delete is not handed out again. Anything holding a stale id gets a
clean miss rather than the wrong slide.

### 5.7 The stale `Slide` object after a delete

- The `Slide` and its part stay alive and readable. Deletion detaches the part
  from the relationship graph; it destroys nothing.
- `slide_index` raises `SlideNotFoundError`; `contains` returns `False`.
- The library **does not poison the object** — no `__class__` reassignment, no
  attribute deletion. It is not our object, python-pptx may hand out the same
  instance again, and mutating someone else's object to enforce a contract
  produces unexplainable bugs later.

### 5.8 Slide part names are not renumbered

A delete leaves a gap (`slide1.xml`, `slide3.xml`). This is benign: content
types are regenerated per-part at save, and relationship targets are computed
relative to the referring part at write time. Nothing depends on the numbering
being dense.

`PresentationPart.rename_slide_parts()` exists upstream and is deliberately
**not** called. It mutates parts the operation did not otherwise touch —
changing the output bytes of untouched slides and making save-to-save diffs
useless for debugging — and it only half-normalizes, renumbering slide parts
but not the notes slides, charts, or diagrams that follow the same convention.

---

## 6. Cross-Deck Copy & Merge

**Status:** *reserved.* Targeted for v0.2; see `ROADMAP.md`. Deliberately
empty so that landing it renumbers nothing.

The one design commitment made now: `layout_strategy` will default to
**`import`** (carry the source layout → master → theme chain into the target)
rather than `remap` (match to the nearest target layout). Fidelity is the
default; `remap` trades visual drift for a cleaner deck and must be asked for.

---

## 7. SmartArt Tiers

**Status:** *reserved.* Preservation is delivered in v0.1 as a consequence of
§4.6, not as an API. Reading, editing, and template-clone authoring are roadmap
items; a layout engine is refused permanently (§1.3).

---

## 8. Fidelity Contract

**Status:** stable.

### 8.1 What survives a part-level clone

Because the copy engine clones parts and rewrites relationships rather than
reconstructing content through python-pptx's object model, content the model
does not understand is carried byte-for-byte:

- **SmartArt** — all five parts (data model, layout definition, quick style,
  colors, and the rendered drawing cache).
- **Embedded video and audio**, and their poster images.
- **OLE objects** and their embedded packages.
- **Custom XML**, tags, and VML fallbacks.
- **Charts** and their embedded workbooks.

This is preservation, not comprehension. `pptx_plus` does not know what a
SmartArt diagram means; it knows the diagram's parts are reachable and that
cloning them faithfully is the whole job.

### 8.2 What does not survive, and why

- **Content routed through python-pptx's authoring API.** Anything a caller
  constructs with `python-pptx` is subject to python-pptx's own fidelity
  limits. This contract covers what `pptx_plus` clones, not what the base
  library writes.
- **A cloned XML part is reparsed**, so ignorable inter-element whitespace is
  normalized. This affects only parts the engine actually rewrites; blob-backed
  parts (media, embedded packages, and every SmartArt definition part except
  the one named below) are copied as bytes and are unaffected.
- **The SmartArt *data* part is the one blob part that is reparsed**, and the
  exception is unavoidable rather than an oversight. It carries
  `dsp:dataModelExt/@relId` naming the rendered-drawing part, so the reference
  has to change on a copy — bytes that must differ cannot also be identical.
  python-pptx models no class for this part, so it loads as an opaque blob
  with no element tree and its bytes must be rewritten *before* `Part.load`;
  a loaded blob part has no public way to change them. See §4.4 on why that
  attribute's scope is the referring part.

### 8.3 What an untouched save must not change

Saving a deck through python-pptx without calling any `pptx_plus` function is
outside this contract — that is python-pptx's behavior. What *is* in the
contract: a `pptx_plus` operation must not alter parts it did not declare it
would touch. Deleting slide 3 must leave slide 1 byte-identical.

---

## 9. Internal Architecture Invariants

**Status:** stable. Numbering aligned with docx-plus `SPEC §9`.

**9.1** No imports between capability packages, and **no relative imports
anywhere**. Every import is absolute, so the AST check in
`tests/test_import_invariant.py` can see it.

**9.2** All element construction goes through `core/oxml.py`'s `el` and `sub`.
No bare `lxml.etree.SubElement` in a capability module.

**9.3** Relationship ids come only from `core/relmap.py` and
`core/parts.py`. A capability module never mints or rewrites an `rId` itself.

**9.4** No magic attributes on python-pptx objects. State lives in the package
or is passed explicitly.

**9.5** All public functions are fully typed; `mypy --strict` passes with zero
ignores.

**9.6** All public functions have Google-convention docstrings naming the
python-pptx gap they fill and citing the relevant ECMA-376 clause.

**9.7** Domain conditions raise typed errors. Every one subclasses
`PptxPlusError`, which lives in `core/errors.py` — a module importing nothing
— so the rest of `core` can subclass it with a plain top-of-file import.
Dual-inherit a stdlib type where it aids `except` ergonomics.

**9.8** No unrequested side effects. In particular, no code path may touch
`SlidePart.notes_slide` or `PresentationPart.notes_master_part` for inspection:
both are lazy properties that *create* parts, the latter creating a notes
master and a theme part. Use `has_notes_slide` and iterate relationships
directly.

**9.9** **The package must round-trip.** An operation touches only the parts it
declares it touches, and leaves the relationship graph closed: every
relationship-id-bearing attribute in every reachable part resolves to a
relationship on that part. This invariant has no docx-plus analog and is the
one this library exists to uphold.

---

## 10. Test Strategy

**Status:** stable. Numbering aligned with docx-plus `SPEC §10`.

**10.1 Layer 1 — structural unit tests.** One assertion per test. The
element is where it belongs, the attribute has the value it should.

**10.2 Layer 2 — round-trip tests.** Every assertion about an operation's
effect runs against a deck **saved to `BytesIO` and reopened**, never against
the in-memory `Presentation`. This is not stylistic: per §3.5 the in-memory
graph keeps stale references while the serialized package is clean, so
asserting in memory tests the wrong artifact.

**10.3 Layer 2.5 — OPC integrity.** The invariant battery, asserted after
every operation:

| | Invariant |
|---|---|
| I1 | Every `p:sldId/@r:id` resolves to a slide relationship on the presentation part |
| I2 | `len(sldIdLst)` equals the count of slide relationships — no orphan rels, no dangling sldIds |
| I3 | Every `p:sldId/@id` is unique and within `[256, 2147483647]` |
| I4 | For every reachable part, every relationship-id attribute resolves (§9.9) |
| I5 | Partnames are unique across the package |
| I6 | Every section entry names a live slide id; every custom-show entry names a live slide relationship |
| I7 | After a duplicate: chart / notes / diagram parts are disjoint from the source's, image and media parts are the *same objects* |

This layer is pure Python, runs on every platform, and is the project's real
gate. It replaces schema-position as the third leg of the triad, because the
correctness condition here is graph-shaped rather than position-shaped.
Schema-position survives as one narrow test on `p:sldIdLst`'s placement.

**10.4 Layer 3 — headless render smoke.** LibreOffice, gated on `soffice`
being on `PATH`, marked `requires_libreoffice`. **A corruption detector, not a
fidelity oracle**: LibreOffice does not run the SmartArt layout engine, so a
diagram slide asserts only that the deck converts cleanly with the expected
page count. Claiming more would be a lie.

Below it, and running everywhere with no external dependency: parse every XML
part of every saved artifact, and verify `[Content_Types].xml` covers every
entry. Those two catch most of what `soffice` would, at zero setup cost.

**10.5 Shared assertion library.** `pptx_plus/_testing/ooxml_asserts.py`,
shipped in the wheel, omitted from coverage, exempt from docstring linting. It
encodes **format invariants**, not test-case expectations.

Before it grades any verb, the harness must itself be graded: run it against
unmodified fixtures to prove it passes on valid input, and against the
deliberately-broken naive recipe (§3.6) to prove it **fails**. A harness that
cannot detect the known-broken recipe cannot grade the correct one.

**10.6 Fixtures.** Generated by `tests/fixtures/build_decks.py` into a
session-scoped temp directory and **never committed**; the builder may not
write into the source tree. The exception is
`tests/fixtures/pptx_samples/`, hand-authored in PowerPoint and committed as
binaries with a provenance README — SmartArt, embedded video, sections, and
custom shows cannot be authored by python-pptx at all.

**10.7 Coverage.** `fail_under = 90`.

---

## 11. Examples

**Status:** deferred to Phase 6.

Runnable via `python -m pptx_plus.examples.<name>`, shipped in the wheel,
omitted from coverage, and **cp1252-safe** — print ASCII only, so they run on
a default Windows console.

---

## 12. Documentation Requirements

**Status:** stable.

MkDocs Material, `strict: true`, four layers: `getting-started`, `guides/`
(task-oriented), `concepts/` (mechanism), `reference/` (mkdocstrings stubs).
The nav grows with the code and never ahead of it — under `--strict`, a nav
entry naming a missing file is a build failure, which is the desired behavior.

**No version literal appears in `README.md`, `SPEC.md`, or anywhere under
`docs/`.** The PyPI badge renders the live version; a hand-typed one only ever
goes stale. Enforced by `tests/test_no_version_literals.py`. Version literals
are confined to `pyproject.toml`, `pptx_plus/__init__.py`, `uv.lock`, CHANGELOG
headings, and ROADMAP's current-state line — all of which the release process
already touches.

---

## 13. Quality Gates

**Status:** stable. Numbering aligned with docx-plus `SPEC §13`.

Every one of these must pass before a commit lands:

```bash
uv run ruff check pptx_plus/ tests/
uv run ruff format --check pptx_plus/ tests/
uv run mypy --strict pptx_plus/
uv run pytest --cov=pptx_plus --cov-fail-under=90
uv run mkdocs build --strict
```

CI adds: the lower-bound dependency leg, the full interpreter matrix including
Windows, the wheel-import smoke, and the LibreOffice tier.

**The manual gate, once per release**, and the real acceptance bar: for each
committed sample deck, duplicate a slide with the library, open the result in
PowerPoint, confirm **no repair prompt**, save from PowerPoint, reopen with the
library, and re-run the invariant battery. It cannot be automated, and nothing
that can be automated substitutes for it.

---

## 14. Build & Packaging

**Status:** stable.

### 14.1 Distribution

hatchling; `uv` for dependency management and every dev command. Runtime
dependencies: `python-pptx>=1.0.2,<2` and `lxml>=4.9`. `py.typed` ships.

Neither Pillow nor XlsxWriter is declared, though python-pptx pulls both:
`pptx_plus` never imports them, because it moves image blobs and embedded
workbooks byte-for-byte and never decodes or parses one. Declaring them would
misstate the dependency.

### 14.2 The upper bound, and `core/_compat.py`

`python-pptx` is capped below 2.0 because this library reaches beneath its
public object model. The surface actually relied upon is small and mostly
public — `Part.load`, `Part.blob`, `Part.relate_to`, `Part.rels`,
`Package.next_partname` — with private use quarantined to two accessors in
`core/oxml.py` (`part_root`, `sld_id_lst`).

`check_upstream_surface()` asserts each depended-on attribute exists at import
time and raises `UpstreamSurfaceError` naming the missing one and the tested
version range. `tests/test_upstream_surface.py` covers each individually. This
converts an upstream break from "a mysterious failure in a user's deck six
months from now" into "one red test naming the attribute, on the day the pin
is bumped."

---

## 15. Post-v0.1 Roadmap

**Status:** pointer, not a list.

`ROADMAP.md` is the single live authority for what is planned, what is
deferred, and what is refused. It is deliberately **not** duplicated here —
that duplication is precisely how a spec and a roadmap drift into disagreeing.

---

## 16. Error Taxonomy

**Status:** stable. Numbering aligned with docx-plus `SPEC §16`.

```
PptxPlusError(Exception)                                # core/errors.py
├── SlideNotFoundError(PptxPlusError, KeyError)         # slides/resolve.py
├── SlideIndexError(PptxPlusError, IndexError)          # slides/resolve.py
├── InvalidNamespaceError(PptxPlusError, ValueError)    # core/ns.py
├── SlideIdRangeError(PptxPlusError, ValueError)        # core/ids.py
├── DanglingRelationshipError(PptxPlusError, ValueError)# core/relmap.py
├── UnclonablePartError(PptxPlusError, TypeError)       # core/parts.py
└── UpstreamSurfaceError(PptxPlusError, RuntimeError)   # core/_compat.py

RelIdLiteralWarning(UserWarning)                        # core/relmap.py
```

Every typed error dual-inherits the stdlib type a caller would naturally catch,
so existing `except KeyError` / `except IndexError` clauses keep working while
`except PptxPlusError` catches everything this library raises.

Plain argument-validation failures inside a function still raise bare
`ValueError` or `TypeError`, with a message that echoes the offending value.

---

## Appendix A: References

- **ECMA-376 Part 1**, 5th edition — §13 (Open Packaging Conventions),
  §19.2.1.34 (`p:sldIdLst`), §19.3.1 (slide), §21.4 (DrawingML Diagrams).
- **python-pptx** — [docs](https://python-pptx.readthedocs.io/),
  [#67 (delete slide)](https://github.com/scanny/python-pptx/issues/67),
  [#132 (duplicate slide)](https://github.com/scanny/python-pptx/issues/132).
- `notes/scope-notes.md` — the discussion artifact this specification
  implements, and the provenance for the scope decisions in §1.
