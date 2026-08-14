# pptx_plus — Implementation Notes

Companion to `SPEC.md`. The spec is the contract: what to build, what the
public API looks like, what the quality gates are. This document is the
meta-guidance: how to think about building it, what order to do it in, where
the implementation will tempt you to cut corners, and how to know when
something is actually working versus merely looking right.

Read both before starting. The spec without these notes leads to reasonable
code that drifts on the things that matter; these notes without the spec lead
to well-organized code that doesn't meet the contract.

Dev-only. Not shipped in the sdist.

---

## 1. Mental Model

The thing to internalize before writing any code: **relationship-graph
correctness is the hard part of this project, and every failure mode is
silent.**

That is a sharper problem than docx-plus's "OOXML correctness is the hard
part." In a `.docx`, a malformed element usually produces something visibly
wrong — Word renders it oddly, or refuses the file. In a `.pptx`, a broken
relationship produces a file that *opens*, looks nearly right, and is missing a
picture. Or opens with a repair prompt three weeks later on someone else's
machine. The XML is well-formed. The schema validates. The graph is wrong.

Four implications follow.

**The oracle comes before the code.** You cannot eyeball a relationship graph.
Build the assertion battery (SPEC §10.3) first, prove it fails on the
known-broken recipe, and only then write the verb it grades. This inverts
docx-plus's phase ordering deliberately. A test suite that cannot detect the
StackOverflow recipe is not a test suite, it is decoration.

**Assert against the saved package, never the in-memory object.** The
in-memory graph keeps stale references by design; the serialized package is the
artifact users actually get. This is the single most important habit in the
repo, and the easiest one to forget when a test is "obviously" fine.

**A test that passes for the wrong reason is worse than no test.** The
canonical example is here in this project: a simple duplicate produces
relationship ids on the clone that happen to be identical to the source's, so
an implementation that skips the id remap entirely passes. The fixture that
catches it (`gap_rids`) has to be constructed on purpose. Whenever a test
passes on the first run, ask what would have to be true for it to pass while
the code is wrong.

**Verification has to end in PowerPoint.** Unit tests catch shape errors and
the integrity battery catches graph errors, but neither catches "PowerPoint
shows a repair prompt." LibreOffice is a corruption detector, not a fidelity
oracle — it does not run the SmartArt layout engine at all. The manual
PowerPoint checklist (SPEC §13) is the real acceptance bar and cannot be
automated away.

---

## 2. Build Order

Ordered by dependency, not importance. Don't parallelize across phases until
Phase 2 is complete.

### Phase 0 — Scaffolding and green CI (0.5–1 day)

Repository skeleton, packaging, CI, prose docs, mkdocs. The goal is a **green
badge on a repository containing no capability code**, so that when CI goes red
during Phase 5 the bisect range is one phase rather than the project.

The smoke test is not `assert True`. It instantiates a `Presentation()` and
asserts the bundled template loads, so every matrix leg proves the dependency
resolved and unzipped rather than merely importing.

**Done when:** all five CI jobs green on `main`, Windows and min-deps legs
included; docs deployed; `pre-commit run --all-files` clean.

### Phase 1 — Foundation (1–2 days)

`core/errors.py`, `ns.py`, `oxml.py`, `_compat.py`, `reltypes.py`, `ids.py`,
`parts.py`. Mostly ported from docx-plus, with the namespace map rewritten for
PresentationML.

Also `tests/test_import_invariant.py`, copied verbatim with
`CAPABILITIES = {"slides"}`. A one-element set looks silly today; it costs
nothing and it is what stops `slides/` from importing `sections/` in v0.3.

**Done when:** the gate is green and a placeholder integration test builds a
deck through `core/` helpers, saves it, and python-pptx reopens it.

### Phase 2 — The OPC integrity harness (0.5 day)

`_testing/ooxml_asserts.py`. Build the oracle before the code it grades.

Grade the harness in three steps, in this order:

1. It **passes** on unmodified generated fixtures.
2. It **passes** on a PowerPoint-authored sample deck.
3. It **fails** on a deck mutated by the naive `_sldIdLst.remove(...)` recipe.

Step 3 is the one that matters and the one that will be skipped under time
pressure. Write it as a real test with `pytest.raises(AssertionError)`, not as
a comment saying you checked.

**Done when:** all three hold.

### Phase 3 — `move_slide`, `resolve.py`, `sections.py` (1 day)

The lowest-risk verb, chosen first because it exercises the entire pipeline —
argument normalization, mutation, save, reopen, assert — while touching no
parts at all. It ships `resolve.py`, which the other two verbs both need, and
`sections.py`, which they both need too.

Pin index semantics in SPEC here (§5.5) and never revisit them. The test that
locks it down is `move_slide(prs, i, i)` parametrized over every valid `i`.

**Done when:** reordering round-trips, the invariant battery passes, and a
sectioned deck's section membership follows the moved slide.

### Phase 4 — `delete_slide` (1–2 days). First genuinely risky work.

The mechanics are three lines. The risk is in what else references the slide.

Roots that are easy to forget:

- The slide's own notes slide — delete it. The notes *master* — keep it.
- Images and media shared with other slides — keep them. This is handled for
  free by reachability, but test it, because "handled for free" is exactly the
  kind of claim that turns out to be false.
- `p14:sectionLst` entries, keyed on **slide id**.
- `p:custShowLst` entries, keyed on **`r:id`**.
- Layouts and masters, which are referenced from `presentation.xml` and must
  survive even when the last slide using them is deleted.

Use `pres_part.rels.pop(rId)`, never `drop_rel` — see the gotcha in CLAUDE.md.

**Done when:** deleting an image-bearing slide where the image is shared leaves
the image present; where it is not shared, the package actually shrinks; the
notes part is gone and the notes master remains; and PowerPoint opens the
result with no repair prompt.

### Phase 5 — `duplicate_slide` and the copy engine (2–3 days). The cliff.

Order within the phase matters:

1. `relmap.py` **standalone**, unit-tested against synthetic maps — including
   the aliasing permutation (`rId1 → rId2` *and* `rId2 → rId1` in one map).
   This is the piece most worth getting right in isolation, because it is pure
   and testable without any package at all.
2. `partgraph.py` — read-only enumeration, so it can be checked against real
   decks before anything mutates one.
3. `clone.py`, then `slides/duplicate.py`.

Fixtures in cost order: `picture` → `gap_rids` → `chart` → `notes` →
`dup_rels` → the committed `smartart` and `video` samples.

**Done when:** a slide carrying an image, a chart, and a SmartArt diagram
duplicates with the battery green; the diagram's parts are byte-identical
clones; two duplicates are independent; and PowerPoint renders the copy
identically to the original.

### Phase 6 — Polish and ship (1–2 days)

Examples, the LibreOffice tier, the property test, docs, the manual checklist,
then tag.

---

## 3. Day 1, Concretely

```bash
cd pptx-plus
uv sync --extra dev
uv run pre-commit install
uv run pytest && uv run mypy && uv run ruff check
```

Then, before writing any capability code, spend twenty minutes in a scratch
script looking at a real package. Not reading about it — looking at it.

```python
from pptx import Presentation

prs = Presentation("some-real-deck.pptx")
for r in prs.part.rels.values():
    print(r.rId, r.reltype.rsplit("/", 1)[-1], r.target_partname if not r.is_external else "EXT")
for sld_id in prs.slides._sldIdLst:
    print(sld_id.id, sld_id.rId)
```

Then `unzip -l` the same deck and compare. The goal is to have the shape of the
graph in your head — which things point at which, and by what kind of id —
before writing code that rearranges it. Everything in SPEC §3 makes sense
immediately once you have seen it and is abstract nonsense until you have.

Use a deck with sections and a chart. A default python-pptx deck will teach you
the easy half only.

---

## 4. Implementation Patterns

**Construct through `core/oxml.py`, never `lxml` directly.**

```python
# Instead of:
#     sld_id = etree.SubElement(lst, "{...}sldId")
#     sld_id.set("{...relationships}id", r_id)
#
# Use:
sld_id = sub(lst, "p:sldId", **{"r:id": r_id})
```

The colon in an OOXML attribute name means the kwargs have to be
dict-splatted. That is ugly and it is the price of one chokepoint.

**Query with XPath variables, never f-strings.**

```python
xpath(root, "./p:sldId[@r:id=$rid]", rid=r_id)  # yes
xpath(root, f"./p:sldId[@r:id='{r_id}']")  # no
```

**Prefer public python-pptx API even when private is shorter.** The clone
primitive is the model here: `type(src).load(partname, ct, pkg, src.blob)` is
entirely public and does more, more reliably, than the `deepcopy(_element)`
approach it replaced. When you find yourself reaching for an underscore,
spend five minutes looking for the public equivalent first — in this codebase
that search has paid off more often than not.

Where private access is genuinely required, quarantine it in one named
accessor in `core/oxml.py` and add a case to `test_upstream_surface.py`.

**Snapshot upstream objects into frozen dataclasses at the boundary.** A
`_Relationship` handed around the codebase spreads the dependency; a frozen
`RelEdge` built from one does not.

---

## 5. The Copy Engine, Concretely

The whole engine is one recursive function with a memo table. If it grows past
roughly a hundred lines, something has become a special case that should have
been a rule.

The three things it must get right, in the order they will bite you:

**Cycle safety.** Write the memo entry *before* recursing. Slide → notesSlide →
slide is a real cycle in any deck with speaker notes, so getting this wrong
does not produce a subtle bug, it hangs. Key the memo on `id(part)`: python-pptx
parts define no `__eq__`, and `iter_parts` relies on identity hashing too.

**Deferred remapping.** Mint all of a part's relationships, then rewrite its
XML. Rewriting as you go means rewriting against a half-built map.

**Reservation.** `next_partname` cannot see parts that exist but are not yet
attached. Two charts on one slide is the minimal reproduction, and it is a
silent overwrite rather than an error.

The clone-map invariant is worth stating out loud because it does more work
than it looks like it does: *within one operation, each source part maps to at
most one destination, and any edge whose target is already mapped resolves to
the mapping.* That single rule handles the notes-slide back-reference,
diamonds, and cycles, with no code specific to any of them. When you find
yourself writing a special case, check whether the invariant already covers it.

---

## 6. Verifying Graph Correctness

Four tiers, cheapest first. Use the cheapest one that can detect the class of
error you are worried about.

1. **The integrity battery** (SPEC §10.3). Pure Python, runs everywhere,
   catches every dangling-reference class of bug. This is where 90% of the
   value is.
2. **Parse every XML part of the saved zip.** Catches well-formedness damage.
3. **LibreOffice headless conversion.** Catches structural problems that lxml
   accepts but a consumer rejects. Not a fidelity check.
4. **PowerPoint, by hand.** The only thing that can tell you the file is
   actually right.

When something is wrong and you don't know why, go to the zip. `unzip -p
out.pptx ppt/slides/_rels/slide2.xml.rels` answers most questions in one
command, and comparing it against the source slide's rels answers most of the
rest.

---

## 7. Common Pitfalls

Collected because each of these has a plausible-looking wrong version.

- **Assuming the relationship-id map is the identity.** It is, in the simplest
  case, which is precisely why this is dangerous.
- **Assuming the map is a bijection.** Collapse makes it many-to-one.
  Bijectivity "sanity checks" fail on valid decks.
- **String-substituting rIds in the serialized XML.** Aliasing makes this wrong
  and the failure is data-dependent.
- **Calling `drop_rel`.** It no-ops when the slide is in a custom show.
- **Touching `.notes_slide` to check whether notes exist.** It creates them.
- **Testing only decks that python-pptx can author.** Those decks have no
  sections, no custom shows, no SmartArt, and no video — i.e. none of the
  things that break.
- **Deriving slide order from part names.** See CLAUDE.md.
- **Building a partname with `os.path.join`.** Backslashes in an OPC part name
  produce a package that fails silently on some consumers and loudly on others.

---

## 8. Test-Writing Guidance

One assertion per test. Name the test after the invariant, not the mechanics:
`test_duplicate_shares_image_part` beats `test_duplicate_2`.

Every operation test follows the same shape:

```python
prs = Presentation(fixture_path)
duplicate_slide(prs, 0)
reopened = roundtrip(prs)  # save to BytesIO, reopen — SPEC §10.2
assert_rel_ids_resolve(reopened)  # the battery
assert ...  # the one thing this test is about
```

For a new fixture, write the test that *fails* first and confirm it fails for
the reason you think. In a project where the failure modes are silent, "it
passes" is weak evidence; "it failed, I changed one thing, it passed" is
strong.

Prefer `io.BytesIO` over `tmp_path` unless the test is specifically about the
saved zip's contents.

---

## 9. When PowerPoint Rejects a File

The repair prompt says nothing useful. The debugging order that works:

1. **Diff the zip listings**, before and after. A missing part or an unexpected
   one localizes it immediately.
2. **Run the integrity battery** on the saved file. If it passes and PowerPoint
   still complains, the battery is missing an invariant — add it, because that
   is a permanent improvement rather than a one-off fix.
3. **Check the side-indexes.** `p14:sectionLst` and `p:custShowLst` are the
   usual culprits and are invisible to python-pptx.
4. **Bisect the deck.** Delete slides until the problem disappears. The last
   one removed is the one to look at.
5. **Compare against PowerPoint's own output.** Do the same operation by hand
   in PowerPoint, save, and diff the two packages. This is slow and it is
   almost always decisive.

---

## 10. Definition of Done, Per Phase

A phase is done when, in addition to its phase-specific criteria in §2:

- The gate is green: `ruff check`, `ruff format --check`, `mypy --strict`,
  `pytest --cov-fail-under=90`, `mkdocs build --strict`.
- Every new public symbol is in its subpackage's `__all__` and has a Google
  docstring citing SPEC and ECMA-376.
- `CHANGELOG.md` has an entry under `## [Unreleased]` explaining *why*.
- Any new invariant is in `_testing/ooxml_asserts.py`, not inline in one test.
- The progress log below has an entry.

---

## 11. What Success Looks Like at v0.1

A user can delete, reorder, and duplicate slides in a real deck — one with
sections, a chart, speaker notes, and a SmartArt diagram — and the result opens
in PowerPoint without a repair prompt, with every picture intact and the
diagram still rendering.

That is a narrow surface and a high bar, and it is the right trade. The
alternative — five verbs that work on decks python-pptx could have authored
itself — would be worth much less.

---

## 12. Progress Log

Newest last. One entry per working session; keep them short and factual.

### 2026-08-13 — Session 1: scope, design, Phase 0

Read `notes/scope-notes.md` and surveyed docx-plus for conventions to mirror.
Settled the two open questions from scope-notes §8: v0.1 ships the floor
(same-deck only), and the eventual merge layout policy defaults to `import`.

Four findings from probing python-pptx 1.0.2 changed the plan:

1. **Orphan collection is free.** `iter_parts` walks the relationship graph and
   `save` writes only what it reaches, so dropping a relationship collects the
   part and its unshared media. Verified against a saved zip's namelist. The
   scope notes' claim that delete "orphans the slide part" is true of the
   in-memory graph and false of the serialized package. `core/` gets no
   collector.
2. **`Part.load` has no subclass overrides anywhere in python-pptx**, so
   `type(src).load(partname, ct, pkg, src.blob)` is a fully public, uniform
   clone primitive. This removed `_element` from the clone path entirely and is
   the largest single piece of de-risking available.
3. **The id remap is mandatory but the obvious test misses it.** `_next_rId`
   fills gaps and `get_or_add` dedupes on `(reltype, target)`, either of which
   shifts ids on the clone — yet the simple case produces an identity map, so a
   skipped remap passes. Hence the `gap_rids` fixture.
4. **`drop_rel` is conditional** and no-ops on a slide that appears in a custom
   show.

Also added to scope: `core/sections.py`. `p14:sectionLst` and `p:custShowLst`
are unmodelled by python-pptx and absent from every generated fixture, so
without it v0.1 could pass its whole suite and still corrupt any deck a user
had organized into sections. ~0.5 day, and not optional.

Phase 0 landed: packaging, CI (with `workflow_call` reuse so the release gate
is the PR gate, Windows as a first-class matrix entry, and a wheel-import
smoke), SPEC, ROADMAP, CHANGELOG, CLAUDE.md, these notes, and the docs
skeleton.

### 2026-08-14 — Session 2: Phases 1–2

Phase 1 landed the foundation: `ns`, `oxml`, `_compat`, `reltypes`, `ids`,
`parts`. Two private-API accessors library-wide, both in `oxml.py`, both
covered by `test_upstream_surface.py`.

Phase 2 landed the integrity harness *before* the verbs it grades, and graded
it in three steps: it passes on every generated fixture, it fails on each naive
recipe from SPEC §3.6 with the specific invariant named, and every individual
assertion has a targeted falsification so none passes vacuously.

Findings from this session:

1. **The harness must read the zip, not the model.** Settled while writing it,
   and it turned out to also be what makes the harness gradeable at all: it can
   be pointed at a package this library never wrote, including deliberately
   corrupted ones built by zip surgery. A model-level assertion could not.
2. **`posixpath.splitext` is wrong for OPC extensions.** `/_rels/.rels` has
   extension `rels`; `splitext` reads the leading dot as a hidden-file marker
   and reports none. Every `.rels` part resolves through the `rels` Default, so
   this made the content-type check fail on every valid deck — caught only
   because step 1 of the grading runs against known-good input first. An oracle
   built after the code it grades would have been "fixed" by weakening it.
3. **`Default Extension="xml"` limits what a content-type check can detect.**
   A slide that loses its specific Override still resolves to
   `application/xml`, so the assertion cannot see it. Stated as a limit in the
   docstring rather than papered over; the falsification test uses an extension
   nobody declares, which is the case that actually occurs.
4. **`<a:hlinkClick r:id="">` has to be authored by hand.** An empty
   relationship id is ordinary — every action-only link is one — but
   python-pptx's hyperlink API only ever mints a relationship, so no generated
   deck contains it. Without the fixture, the harness would have reported every
   real deck's action buttons as dangling references.

`UNQUALIFIED_REL_ID_ATTRS` moved from the planned `relmap.py` to `reltypes.py`:
the harness needs it three phases before the rewriter exists, and it is a
constant with two consumers. SPEC §4.3/§4.4 updated.

### 2026-08-14 — Session 2 (cont.): Phase 3

`slides/resolve.py`, `slides/move.py`, `core/sections.py`. 270 tests, 99.29%.

One design question actually had to be settled here, and the SPEC did not
answer it: **when a slide moves across a section boundary, what happens to the
sections?**

Two models produce the same running order and disagree about membership. The
*fixed-boundary* model keeps each section's size and leaves the cut at a fixed
position; the *moved-slide* model relocates exactly the moved slide, shrinking
its source section and growing its destination.

Fixed-boundary is what "sections partition the deck" suggests on first reading,
and it is wrong: moving slide 1 into section 2 drags slide 3 back into
section 1. A bystander changes section because of a call that never named it.
Implemented the moved-slide model and asserted the property directly —
`test_reordering_changes_no_other_slides_section` computes the full
id-to-section map before and after and requires the diff to be exactly the one
slide.

The related ambiguity is landing *on* a boundary, where appending to one
section and prepending to the next are indistinguishable in the running order.
The tie goes to the section the slide came from, which is what makes
`move_slide(prs, i, i)` a no-op for the sections as well as for the order —
and no-op-ness is the SPEC's own acceptance criterion for the index semantics
(§5.5), so the two now rest on the same property.

Both are recorded in SPEC §4.7, which is new.

Smaller notes:

- `sections` and `custom_show` moved out of the integrity test's local helpers
  and into `build_decks.py`. Three phases need them now, and a structure that
  exists only inside one test file is a structure the other tests will quietly
  do without.
- `resolve_slide` matches a `Slide` by the identity of its `<p:sld>` element.
  python-pptx defines no `__eq__` on `Slide`, so equality is identity anyway —
  but going through the element says what is meant and survives an upstream
  change to how wrapper objects are cached. Added `Slide.element` to
  `REQUIRED_SURFACE`.

### 2026-08-14 — Session 2 (cont.): Phase 4

`slides/delete.py`. 319 tests, coverage holding at 99%.

Less eventful than budgeted, and the reason is finding 1 from session 1:
because `save` serializes a reachability walk, "collect the orphaned parts" is
not a step. Dropping the presentation part's relationship *is* the collection,
and the notes slide and any unshared image go with it at no cost. The verb is
four statements.

What needed care was ordering and the choice of removal call:

- **Scrub the side-indexes first**, while both identifiers naming the slide
  still resolve. Sections key on `sldId/@id`, custom shows on `@r:id`.
- **`rels.pop`, not `XmlPart.drop_rel`.** Worth recording precisely, because
  the comment in the code makes a claim that looks like over-caution: after the
  scrub there is only one `@r:id` reference left, so `drop_rel` *would* work.
  It is still wrong to use it — correctness would then depend on the ordering
  of two statements in this function rather than on the removal doing what its
  name says. `test_a_slide_in_a_custom_show_really_is_referenced_twice` pins
  the premise so the reasoning stays checkable.

Tests worth keeping honest about: `test_the_package_actually_shrinks` asserts
on `len(blob)` rather than on part absence, because "the part is not in the
namelist" is what the naive recipe's *defenders* would also expect to see if
they only checked the object graph.

Added a `slide_jump` fixture and two tests for the documented v0.1 limitation:
slide B linking to slide A holds an `RT.SLIDE` rel to A's part, so deleting A
leaves that part reachable and still written, alive with no `p:sldId`. The
package remains *valid* — the battery passes — which is exactly why the
limitation needs its own explicit test rather than trusting the battery to
notice. `scrub_links=True` is the v0.2 fix.

Still outstanding: PowerPoint has not been running this session, so the manual
acceptance pass (open each result, confirm no repair prompt) and the
pptlive-authored `pptx_samples/` are both unstarted.

### 2026-08-14 — Session 2 (cont.): Phase 5

`core/relmap.py`, `core/partgraph.py`, `core/clone.py`, `slides/duplicate.py`.
453 tests, 99.38%. Budgeted as "the cliff"; it was the smoothest phase, because
Phase 1's primitives and Phase 2's oracle had already absorbed the risk. The
duplicate tests passed on the first run.

**A category the plan did not have.** `RT.SLIDE` is neither shareable nor
structural nor deep. A slide-jump hyperlink is a relationship to *another
slide*, so deep-cloning it would make duplicating slide 2 silently produce a
second copy of slide 1 as well. Added `REUSE_RELTYPES` and a `Disposition.REUSE`
rather than widening `STRUCTURAL_RELTYPES`: both mean "reuse the target" today,
but structural relationships are the cross-deck seam where the layout chain gets
*imported*, and a slide-jump target has entirely different cross-deck semantics.
Merging them would have handed §6 the wrong behaviour silently.

The notes back-reference is *not* this case, and the distinction is worth
keeping straight: notesSlide -> slide is a cycle, not a reference to a third
party. The clone map already resolves it, which is why the map is consulted
*before* the disposition rather than after.

**The `to_index` default changed from the plan.** Scope-notes said "appends
unless to_index given"; asked, and the answer was adjacent placement, matching
PowerPoint's Duplicate Slide. `to_index=-1` appends. SPEC §5.4 updated.

**The coverage gate broke for a reason unrelated to the tests.** Enabling
`filterwarnings = ["error::pptx_plus.core.relmap.RelIdLiteralWarning"]` in
`pyproject.toml` — the TODO left in Phase 0 — dropped total coverage from 99%
to 60%, with whole modules reading 0%. Pytest resolves an ini filter by
importing the named class at config time, which imports `pptx_plus` before
coverage starts, so every module-level statement records as unexecuted. Moved
the registration to `pytest_configure` in `tests/conftest.py`, where the
warnings plugin parses it per test item long after coverage is running. Worth
recording because the symptom points nowhere near the cause: nothing about a
warning filter suggests a coverage number.

Added `test_the_warning_is_fatal_in_this_suite`, which fails if the promotion
is ever lost. `pytest.warns` passes with or without the filter, so every other
warning test would have kept passing.

Left uncovered deliberately: the early-return in `clone.py`'s `_clone`. It is
unreachable through the current call structure — the edge loop consults the map
before recursing — but it is the termination guard for a deep cycle, and a
`pragma: no cover` on a real safety branch reads worse than one uncovered line.

### 2026-08-14 — Session 2 (cont.): Phase 6, and what the samples caught

Property tests, the runnable example, the LibreOffice tier, the API reference,
and the committed PowerPoint samples. 499 tests, 99%+.

**The samples paid for themselves on the first run**, which is the entry worth
keeping. Authored through `pptlive` (Tom's COM driver) rather than by hand, so
they are reproducible and the click-path is a script rather than a memory.

`smartart.pptx` exposed two defects, both invisible to every generated fixture:

1. **The harness rejected valid PowerPoint output.** `dsp:dataModelExt/@relId`
   resolves against the relationships of the part that *refers* to the diagram
   data part — which has none of its own. `assert_rel_ids_resolve` checked it
   against the containing part and reported a pristine, PowerPoint-authored
   deck as damaged. Step 1 of the SPEC §10.5 grading procedure ("passes on
   valid input") is what caught it, and it only caught it once the valid input
   was something PowerPoint made.
2. **`duplicate_slide` silently produced a wrong deck.** python-pptx has no
   model class for a diagram data part, so it loads as an opaque blob with no
   element tree, and `_is_xml` was False — the rewrite pass skipped it
   entirely. The copy kept the source's `relId` and its diagram pointed at the
   *original's* drawing cache. No exception, no repair prompt: PowerPoint
   recomputes the drawing on open and looks fine.

Defect 2 is the exact failure mode this library exists to prevent, reproduced
in the library itself. It survived Phase 5 because the only SmartArt in the
project until now was a `dsp:dataModelExt` element I constructed by hand in a
unit test — which had the attribute but not the *scope*, since a synthetic
element has no surrounding package to be wrong about.

The fix has three parts: the registry records a scope per attribute;
`remap_rel_ids` takes a `parent` map; and the engine defers cloning
parent-scoped parts until the referring part's map is complete, rewriting
their bytes before `Part.load` because a loaded blob part has no public way to
change them. That makes the data part the one blob part this library reparses
— unavoidable, since the reference has to change, and now stated in SPEC §8.2
rather than left as an unadvertised exception.

Manual acceptance ran for real, not as a checklist item: duplicated the
SmartArt slide, opened in PowerPoint (no repair prompt), rendered both slides
with `pptlive snapshot` and compared them, saved from PowerPoint, reopened,
re-ran the battery, and duplicated again from PowerPoint's own output.
PowerPoint's re-save kept two independent five-part diagram sets, which is the
real confirmation the copy was not an alias.

Two smaller things:

- **Enabling the `RelIdLiteralWarning` filter in `pyproject.toml` broke the
  coverage gate**, dropping the total from 99% to 60% with whole modules at
  0%. Pytest resolves an ini `filterwarnings` entry by importing the named
  class at config time, before coverage starts. Moved to `pytest_configure`.
  The symptom points nowhere near the cause.
- **`ruff format` rewrites Python inside Markdown fences**, so a bare
  `ruff format` was silently editing `SPEC.md` and `notes/`. `*.md` excluded.

Still open: no embedded-video sample (pptlive has no media verb, so it needs a
manual Insert > Video), and the `custom_show` fixture remains hand-built and
unverified against PowerPoint's own shape.
