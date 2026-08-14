# pptx_plus — Roadmap

The single authoritative roadmap for `pptx_plus`. `SPEC.md §15` defers to this
file, and this file does not restate the API contract — that is SPEC's job.

`pptx_plus` is, and stays, a lean extension to `python-pptx` that does the
things `python-pptx` can't. Every item below either fills a documented
`python-pptx` gap or rounds out a surface already started here. Ideas that
don't fit that charter are routed to sibling projects, not absorbed.

Scope decisions here trace back to [`notes/scope-notes.md`](notes/scope-notes.md),
the discussion artifact that preceded the repository.

## Current state — v0.1.0 in development

No releases yet. `core/` (less `relmap`, `partgraph`, `clone`, `sections`) and
`_testing/` are implemented; `slides/` is next. The v0.1 target:

| Module | Surface |
|---|---|
| `core/` | Namespaces, element construction, relationship-id remapping, part-graph cloning, partname and slide-id allocation, section and custom-show maintenance |
| `slides/` | `delete_slide`, `move_slide`, `duplicate_slide` — same deck |
| `_testing/` | The OPC integrity assertion battery (SPEC §10.3), shipped in the wheel |

**The v0.1 thesis is one coherent capability, not a broad surface.** The
slide lifecycle is python-pptx's largest and most-requested gap — deletion and
duplication are its two oldest open feature requests, at roughly twelve years
each — and no open-source library does it correctly today. Shipping that one
thing properly is worth more than shipping five things approximately.

## v0.2 — cross-deck copy and merge

The natural next layer, and the one the `core/` engine was deliberately shaped
around: `clone_part_graph` already takes an `into: Package` parameter that v0.1
always sets to the source's own package.

- `copy_slide(slide, *, into, to_index=None, layout=None)` — imports image and
  chart blobs into the target and resolves the layout → master → theme chain.
- `merge_presentations(target, *sources, layout_strategy="import")`.

**`layout_strategy` defaults to `import`** — carry the source's layout chain
into the target. It grows the deck but never drifts visually. `remap` (match to
the nearest target layout by name and type) produces a cleaner deck at the cost
of silent visual drift, which is the wrong default for a library whose thesis
is fidelity. Decided up front; recorded in `SPEC.md §6`.

Layout and master reconciliation is the genuinely hard part here — the merge
equivalent of what relationship rewriting is to duplication. It is the least
certain estimate in the project, which is exactly why it is not in v0.1.

Also likely in this cycle:

- **`scrub_links=True`** for `delete_slide`, removing slide-jump hyperlinks
  that pointed at the deleted slide. See *Known limitations* below.
- **A `pptx-plus` CLI.** Held out of v0.1 because the three verbs mutate a deck
  in place and need `--output` / backup semantics before a shell should be
  allowed near them.
- **A packaged agent skill.** docx-plus added its equivalent at v0.5, but the
  case is stronger here: the stated primary audience for `pptx_plus` *is* LLM
  tooling. It still should not land before the surface it describes is stable.

## v0.3+ — SmartArt, in tiers

SmartArt is not a binary. A diagram is a `<p:graphicFrame>` referencing four
definition parts — data model, layout definition, quick style, colors — plus a
fifth **rendered drawing cache** holding the positioned shapes every viewer
actually draws.

| Capability | Status | Mechanism |
|---|---|---|
| **Preserve** across duplicate | **v0.1, free** | Part-level clone keeps all five parts; the model is never in the loop |
| **Read** node text and hierarchy | roadmap | Walk the data model's point and connection lists |
| **Edit** node text in place | roadmap | Rewrite text runs in the data part |
| **Author** from a content tree | roadmap, with caveats | Clone a stock layout/style/colors verbatim, rewrite the data model, invalidate the cache, let PowerPoint render |
| Self-compute geometry | **never** | See *Considered, not on the roadmap* |

The loophole that makes authoring tractable at all: **PowerPoint regenerates
the drawing cache on open** from the data model plus the layout. So a correct
data model referencing a *stock* layout renders correctly once opened in
PowerPoint, without anyone computing geometry.

Three caveats that must be documented before any authoring API ships:

- **PowerPoint-render-target only.** Until opened and saved in PowerPoint, the
  diagram is blank in LibreOffice, in web and thumbnail viewers, and in
  headless PDF export — none of them run the layout engine. Fine when the
  output is consumed by opening it in PowerPoint; not fine for a headless
  export pipeline.
- **Layouts constrain structure.** Built-in layouts expect particular node
  counts and nesting depths. Support the diagram types whose constraints can be
  satisfied, not "any tree into any layout."
- **Stale-cache risk.** Cloning a template's cache instead of invalidating it
  makes PowerPoint show the template's old text until it recomputes. Correct
  invalidation is the fiddly part.

If authoring lands, it mirrors docx-plus's "defaults extracted from real
Word-saved samples, not guessed": ship a small catalog of real
PowerPoint-authored diagrams and rewrite only the data model. The catalog needs
to cover the five or so diagram types actually used in practice — list,
process, cycle, hierarchy, relationship — not the full built-in set.

## Backlog — bounded, unscheduled

Each of these is a real python-pptx gap, small enough to scope, and waiting on
demand rather than on design.

- **Slide sections as a first-class API.** v0.1 *maintains* `p14:sectionLst`
  because not doing so corrupts decks; it does not let you create, rename, or
  reorder sections. That is a natural follow-on.
- **Comments.** PowerPoint's modern comment format is a separate part python-pptx
  does not model.
- **Theme and color-scheme helpers.**
- **Hyperlink and table-style helpers.**
- **`renumber_slide_parts()`** as an explicit opt-in. Deliberately never
  automatic — see `SPEC.md §5.8`.

## Considered, not on the roadmap

Refusing things clearly is half of what this document is for.

- **Animations and transitions.** An XML rat's nest with no good abstraction,
  and open upstream since 2018. A library that half-supports animations is
  worse than one that doesn't: the failure mode is a deck that opens fine and
  animates wrong, which is very expensive to debug and impossible to test
  headlessly.
- **A SmartArt layout engine.** The layout definition part is a declarative
  layout *program* — algorithm nodes, constraint lists, `forEach`,
  `choose`/`if` — that PowerPoint executes to compute geometry. No open
  implementation exists; LibreOffice's partial import is famously imperfect.
  Self-computing geometry is the only route to arbitrary custom diagrams and
  guaranteed cross-renderer rendering, and it is refused permanently.
- **Templating.** A different problem, well served elsewhere.
- **A `Presentation` subclass** or any wrapper replacing python-pptx's object
  model. Callers keep their own object; that is the whole compositional
  premise.
- **Live PowerPoint automation (COM).** Different problem, different platform
  constraints, different failure modes.
- **Feature parity with the commercial libraries.** The bar is "useful and
  maintainable by one person," not "market-beating."

## Known limitations, tracked rather than fixed

- **Slide-jump hyperlinks across a delete.** A slide whose hyperlink targets
  another slide holds a relationship to that slide's part. Deleting the target
  leaves its part alive — still reachable from the linking slide — with no
  entry in `sldIdLst`. PowerPoint tolerates this, but the link goes nowhere;
  real PowerPoint removes the hyperlink. Detected and marked `xfail` in v0.1;
  `scrub_links=True` in v0.2.
- **A slide whose layout is not in the deck's master.** Legal in decks
  assembled by other tools. A duplicate inherits the same anomaly, which is
  correct behavior — noted so it is not mistaken for a bug introduced here.
- **Not thread-safe** against concurrent mutation of one package. Neither is
  python-pptx.
