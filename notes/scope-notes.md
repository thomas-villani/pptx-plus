# pptx-plus v0.1 scope — discussion artifact

Discussion artifact, not a spec. Lives in the docx-plus repo only because
no `pptx-plus` repo exists yet; move it when one does. Decides what (if
anything) a `pptx-plus` v0.1 should ship, mirroring the `docx-plus`
"does what the base library can't" thesis.

Sibling-repo per the established pattern (docx-plus / wordlive / …):
`pptx-plus` is a lean composable extension over `python-pptx`, not a
fork and not a feature-parity play against the commercial libs.

**Primary use case (decided 2026-05-21):** backing the author's own
LLM tooling for generating / templating / manipulating decks, released
openly for anyone. The bar is "useful and maintainable by one person,"
not "market-beating." Round-trip fidelity on real decks matters more
than breadth.

## 1. Empirical findings

### 1.1 The central gap — slide lifecycle (verified 2026-05-21)

`python-pptx` has no API to delete, duplicate, reorder, move, or copy
slides. The two requests are among the oldest open issues in the project
(delete = #67, duplicate = #132) and have sat on the backlog for ~12
years. This is the most-asked-for missing capability, by a wide margin.

Why the workarounds are wrong, not just absent:

- The common `deepcopy(slide.element)` + append-to-`sldIdLst` recipe
  copies the slide's `spTree` XML but **leaves every `r:id` / `r:embed`
  pointing at the source slide's relationships.** Images, charts, and
  media are relationship-targeted parts; the copy produces dangling
  references and the picture/chart silently drops on open.
- The common delete recipe (`prs.slides._sldIdLst.remove(...)`) unlinks
  the `<p:sldId>` but **orphans the slide part and its rels** — the
  bytes stay in the package (bloat) and the rels are never cleaned.

So the gap isn't "no convenience method"; it's that **relationship and
part rewriting is the actual hard part, and every naive recipe skips
it** — structurally the same situation as docx-plus comments, where
"range anchoring is the hard part and python-docx skips it entirely."

### 1.2 OOXML mechanics (the work to do correctly)

- Slide order lives in `presentation.xml` → `<p:sldIdLst>` →
  ordered `<p:sldId id=... r:id=...>` children. Reorder = reorder those
  children. Delete = remove child **+ drop the package relationship +
  remove the orphaned slide part (and its now-unreferenced sub-parts)**.
- Duplicate (same deck) = new slide part whose `spTree` is a faithful
  clone, **with every `r:embed`/`r:id` remapped to fresh relationships
  on the new part**, then a new `<p:sldId>`. Image parts can be shared
  by reference; chart parts carry their own embedded `.xlsx` workbook
  sub-part that must be cloned too.
- Copy-between-decks / merge = the above, but the target rels live in a
  **different package**: image blobs must be added to the target,
  chart + workbook sub-parts cloned across, and — the deep part — the
  copied slide's **layout → master → theme chain** must be either
  imported into the target or remapped onto an existing target layout.
  Layout/master reconciliation is the merge equivalent of comments'
  range anchoring: the part that makes it genuinely hard.

### 1.3 A fidelity upside worth noting

If pptx-plus copies slides at the **part + relationship level** (clone
the slide part's XML and its related parts byte-faithfully) rather than
reconstructing through python-pptx's lossy object model, then SmartArt /
embedded video / custom XML on the *copied* slide **survive** — because
they never round-trip through the model that strips them. This sidesteps
python-pptx's known re-serialization data loss for the copy path
specifically. Real differentiator; design the copy engine this way.
(See §6 — this is what makes SmartArt *preservation* free in v0.1.)

### 1.4 Base-library maintenance status (verified 2026-05-21)

The fork's "python-pptx is no longer maintained" pitch is stale:

| | upstream `python-pptx` | `python-pptx-ng` fork |
|---|---|---|
| Latest release | **1.0.2 (early 2025)**, 1.0.0 mid-2024 | **0.7.0 (Jan 2024)** |
| License | MIT | MIT |
| Slide management | No | No (didn't add it) |
| SmartArt | No | No (explicit non-goal) |
| Round-trip fidelity fix | No | No |
| Adds | — | charts (XY/bubble/radar/area/doughnut), gradient fills, table cell merge/split |

Upstream reached 1.0 and shipped patches into 2025 (3.4k★, 444 open
issues); the fork's last release now *predates* upstream's 1.0.0. The
fork is the staler tree today, and it solves none of the gaps that
motivate pptx-plus. This resolves the base-library question — see §2.

## 2. Architecture stance (decided 2026-05-21)

1. **Base = compose with upstream `python-pptx` (1.0.x).** Not the
   `-ng` fork: it is staler than upstream, addresses none of our gaps,
   and occupies the same `pptx` import namespace (either/or swap, not a
   layer). It's MIT, so its chart code can be cherry-picked later if a
   need appears. Not a hard fork: see (4).
2. **Slide-lifecycle ops operate at the OPC layer** — the package,
   parts, relationships, and raw part XML via lxml — **below**
   python-pptx's high-level object model. python-pptx is used only as a
   thin OPC reader/writer for these ops; its lossy re-serialization is
   never in the loop, so there is nothing to "fight." For ordinary
   content (text / tables / pictures / basic charts) use python-pptx's
   existing authoring surface as-is.
3. **"Fighting the base" only arises when asking python-pptx to author
   or model content it doesn't understand** (SmartArt authoring,
   animations). v0.1 avoids that entirely; SmartArt is approached via
   part-level / template-clone strategies that route *around* the model
   (§6), not through it.
4. **The fork decision is deliberately deferred.** Revisit only on
   concrete evidence of a wall that requires fixing model internals —
   e.g. byte-safe round-trip of SmartArt / video / customXml *through*
   the object model. The OPC-layer copy engine already delivers
   copy/merge fidelity without owning the model, so a fork is unlikely
   to be justified for the LLM-tooling use case. Architect so this
   stays a late, evidence-driven choice rather than an upfront tax.

## 3. What pptx-plus would add for v0.1

v0.1 thesis: **slide lifecycle, done correctly.** One coherent capability
with no good open-source equivalent. Everything additive (comments,
sections, theme helpers) is explicitly v0.2; SmartArt beyond
preservation is roadmap (§6).

### 3.1 `core/` — the relationship-aware copy engine

The foundation everything else composes with. Not user-facing on its
own. Operates at the OPC layer (§2.2), not through python-pptx's model.

- Part/relationship graph walk for a slide: enumerate the slide part and
  every part reachable by relationship (images, charts + their workbook
  sub-parts, media, and — preserved, not understood — diagram parts).
- `clone_part_graph(slide_part, into_package) -> new_slide_part`:
  faithful byte-level part clone + fresh relationship minting + `r:id`
  remap inside the cloned XML. Idempotent rel-id allocation.
- Orphan collection: after a delete, drop parts no longer referenced
  from any `sldId`/layout/master.

This is the pptx analog of `docx_plus/core/` (oxml builders, NSMAP,
schema-strict insertion, parts management). Reuse those patterns:
no raw lxml in capability modules, namespaces from one `ns.py`,
round-trip + idempotency + schema-position tests with shared helpers.

### 3.2 `slides/` — the public verbs

```python
def delete_slide(prs, slide_or_index) -> None:
    """Unlink the sldId, drop the relationship, collect the orphaned
    slide part and any now-unreferenced sub-parts. Idempotent."""

def move_slide(prs, slide_or_index, to_index: int) -> None:
    """Reorder within sldIdLst. Pure ordering; no part changes."""

def duplicate_slide(prs, slide_or_index, *, to_index: int | None = None) -> Slide:
    """Clone within the same deck via the copy engine (rel/embed remap).
    Appends unless to_index given. Images shared by ref; charts +
    workbooks cloned. SmartArt/video preserved (part-level clone)."""

def copy_slide(slide, *, into: Presentation, to_index: int | None = None,
               layout: SlideLayout | str | None = None) -> Slide:
    """Cross-deck copy. Imports image/chart blobs into `into`. Resolves
    the layout/master chain: reuse `layout` if given, else import the
    source chain. Part-level clone preserves SmartArt/video on the slide."""

def merge_presentations(target, *sources, layout_strategy="import") -> None:
    """Append all slides from each source into target via copy_slide.
    layout_strategy: 'import' (carry source layouts/masters) or
    'remap' (match to nearest target layout by name/type)."""
```

| Verb | Difficulty | Notes |
|---|---|---|
| `move_slide` / reorder | Low | `sldIdLst` child reordering only |
| `delete_slide` | Low–Med | unlink + rel drop + orphan collection |
| `duplicate_slide` (same deck) | Med | copy engine; chart workbook sub-part is the fiddly bit |
| `copy_slide` (cross-deck) | High | blob import + layout/master chain |
| `merge_presentations` | High | repeated copy_slide + layout reconciliation |

## 4. Cost picture

| Module | Days | Confidence |
|---|---|---|
| `core/` copy engine (part graph + rel remap + orphan collection) | 3–5 | Medium |
| `slides/` delete + move + reorder | 1–2 | High |
| `slides/` duplicate (same deck) | 2–3 | Medium |
| `slides/` copy-between-decks | 3–4 | Medium–Low |
| `slides/` merge + layout/master reconciliation | 3–4 | Low |
| Test harness (round-trip, LibreOffice smoke, idempotency) | ~2 | High |

- **Floor (delete + move + duplicate, same-deck only)**: core + low/med
  verbs + tests ≈ **8–12 days**. Already a shippable, defensible v0.1 —
  it closes #67 and #132, the two headline gaps.
- **Full (incl. cross-deck copy + merge)**: ≈ **14–19 days**, with the
  merge tail the least certain (layout/master reconciliation is open-
  ended). Strong candidate to defer cross-deck/merge to v0.2.

## 5. Decision factors

Arguments to build:

- The flagship gap is larger and more universal than any single Word
  gap — a verb everyone reaches for, with visibly broken alternatives.
- The part-level copy engine *preserves* SmartArt/video on copied
  slides, turning python-pptx's worst limitation into a selling point
  for the copy path.
- The engineering identity (fiddly OOXML done correctly, idempotent,
  round-trip-tested) is exactly what docx-plus already proved you can
  execute, and the core/ patterns transfer conceptually.
- Fits the sibling-repo philosophy cleanly.

Arguments to hold:

- The features that get the most *attention* (animations, SmartArt
  authoring) are the ones we must scope carefully or refuse.
- Same-deck duplicate is the only "easy win"; the headline cross-deck
  merge is the riskiest item in the table.

## 6. SmartArt — tiered capability, not a binary

SmartArt is a `<p:graphicFrame>` referencing **four definition parts**
(data model `dgm:dataModel`, layout `dgm:layoutDef`, quickStyle, colors)
plus a fifth **rendered drawing cache** (`dsp:` namespace) that holds the
actual positioned shapes — what every viewer draws.

The wall is the **layout part**: it is not a static template but a
declarative layout *program* (algorithm nodes, constraint lists,
`forEach`, `choose`/`if`). PowerPoint runs it against the data model to
compute the drawing cache. Reimplementing that engine (the only way to
self-compute geometry / support arbitrary custom layouts / guarantee
cross-renderer rendering) is the genuine non-goal — no open
implementation exists; LibreOffice's partial import is famously
imperfect.

The loophole: **PowerPoint regenerates the cache on open** from data +
layout. So a correct data model that references a *stock* layout/style/
colors renders correctly once opened in PowerPoint, without us computing
geometry. That yields a capability ladder:

| Capability | Tier | Feasibility | Mechanism |
|---|---|---|---|
| **Preserve** SmartArt across copy/duplicate/merge | **v0.1 (free)** | High | OPC-layer part clone keeps all 5 parts; never touches the model (§1.3) |
| **Read** node text + hierarchy | v0.2 | Easy | walk `dgm:ptLst` / `dgm:cxnLst`, extract `<a:t>` |
| **Edit** node text in place | v0.2 | Achievable | rewrite text runs in the data part (known op; Aspose does this) |
| **Author** from a content tree (template-clone) | v0.2–v0.3 | Achievable w/ caveats | clone stock layout/style/colors verbatim, rewrite data model, invalidate cache, let PowerPoint render |
| Self-compute layout / arbitrary custom diagrams / cross-renderer fidelity | **never** | Not feasible | requires the layout engine |

**Template-clone authoring** mirrors docx-plus's "defaults extracted
from real Word-saved samples, not guessed": ship a small catalog of
real PowerPoint-saved diagrams, clone their definition parts, and
rewrite only the data model. The catalog need cover **only the ~5–10
diagram types actually used in practice** (list, process, cycle,
hierarchy/org, relationship) — not the full built-in set, most of which
see little real use. A natural LLM-facing API:
`add_smartart(slide, layout="process", nodes=[...tree...])`.

Authoring caveats to document up front:

- **PowerPoint-render-target only.** Until opened-and-saved in
  PowerPoint, the diagram is blank in LibreOffice, web/thumbnail
  viewers, and headless PDF export (none run the engine). Fine if output
  is consumed by opening in PowerPoint; not fine for a headless PDF step.
- **Layouts constrain structure.** Built-in layouts expect particular
  node counts/levels; support the diagram types whose constraints we can
  satisfy, not "any tree into any layout."
- **Stale-cache risk.** If the template's cache is cloned instead of
  invalidated, PowerPoint may show the template's old text until it
  recomputes. Correct cache invalidation is the fiddly bit.

## 7. Out of scope for v0.1 (document up front)

- Animations / transitions (XML rat's nest, open upstream since 2018).
- SmartArt beyond *preservation* — read / edit-text / template-author
  are roadmap per §6, not v0.1.
- Full chart editing beyond what survives a part-level clone.
- Comments, slide sections, theme/color helpers, hyperlink/table-style
  helpers — the additive v0.2 docx-plus analogs, deferred to keep v0.1
  a single coherent capability.

## 8. Open questions for the user

1. **v0.1 line**: ship the floor (same-deck delete/move/duplicate) and
   queue cross-deck copy/merge for v0.2, or commit to the full set up
   front despite the merge tail's low confidence?
2. **Merge layout policy default**: `import` (carry source
   layouts/masters — fidelity, deck bloat) vs `remap` (match to target
   layouts — clean deck, possible visual drift). Which is the sane
   default?

*(Resolved: base library = compose with upstream python-pptx, slide-ops
at the OPC layer — §2. Audience = author's own openly-released LLM
tooling — see intro.)*
