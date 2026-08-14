# CLAUDE.md

Guidance for AI coding agents (and humans) working in this repository.

## What this is

`pptx_plus` is an OPC-level extension layer for
[python-pptx](https://github.com/scanny/python-pptx). It reaches the part of
the `.pptx` format python-pptx does not expose — the slide lifecycle: deleting,
reordering, and duplicating slides with the relationship graph left closed —
while leaving the underlying `Presentation` object fully usable.

- **Scope discipline:** keep this a lean python-pptx *extension*. It is not a
  deck-authoring framework and does not do live PowerPoint automation.
  Adjacent ideas belong in sibling projects, not here.
- Authoritative API contract: `SPEC.md`. Live status: `ROADMAP.md`.
  Per-release record: `CHANGELOG.md`. Code cites SPEC; nothing cites ROADMAP.
- Development meta-guidance — build order, how to know something actually
  works, where the implementation will tempt you to cut corners:
  `IMPLEMENTATION.md`.

## Environment & tooling

This project uses **`uv`** for everything. Never call bare `python` or `pip`.

```bash
uv sync --extra dev          # install package + dev deps (single source: [project.optional-dependencies] dev)
uv run pytest                # run the test suite
uv run pytest tests/test_foo.py -k name   # one file / one test
uv run mypy                  # strict type-check (files = ["pptx_plus"])
uv run ruff check            # lint  (rules: E,F,W,I,B,UP,D — Google docstrings)
uv run ruff format           # format (line-length 100)
uv run mkdocs serve          # preview docs locally
uv run mkdocs build --strict # docs must build link-clean
```

Pre-commit mirrors the CI lint gate: `uv run pre-commit run --all-files`.

## Architecture

Layered, one-way dependencies:

- `core/` — foundation: `PptxPlusError` (base of every typed error), the
  namespace map (`ns`), OOXML element helpers (`oxml`), relationship-type
  constants (`reltypes`), id allocation (`ids`), part allocation and cloning
  (`parts`), the relationship-id rewriter (`relmap`), part-graph enumeration
  (`partgraph`), the clone orchestrator (`clone`), section and custom-show
  maintenance (`sections`), and the upstream-surface guard (`_compat`).
  Depends on nothing above it.
- **Capability packages** — `slides/` at v0.1. Each builds on `core/` and is
  independent of its siblings.
- `_testing/` — test-only OPC assertions; shipped in the wheel, excluded from
  coverage and the public API.

Each subpackage's `__init__.py` `__all__` is the authoritative public surface
for that package.

## Conventions

- **Errors:** every public exception subclasses `core.PptxPlusError` and
  dual-inherits the stdlib type a caller would naturally catch (e.g.
  `SlideNotFoundError(PptxPlusError, KeyError)`). SPEC §16.
- **Typing:** `mypy --strict` must pass with zero ignores. `from __future__
  import annotations` in every module; `if TYPE_CHECKING:` for python-pptx
  types. Keyword-only arguments after `*` on public functions.
- **Docstrings:** Google convention. Tests, `_testing/`, and `examples/` are
  exempt. A module docstring names the python-pptx gap it fills, cites the
  ECMA-376 clause, and ends with the invariant line
  ``This module imports only from ``pptx_plus.core`` (SPEC §9.1).``
- **Python:** target 3.10+. CI tests 3.10–3.13 on Linux plus 3.13 on Windows.
- **Coverage:** `fail_under = 90`. New code needs tests.
- **No version literals** in `README.md`, `SPEC.md`, or `docs/**` — enforced by
  `tests/test_no_version_literals.py`. The PyPI badge renders the live version.

## Gotchas

Hard-won, and every one of them has bitten a naive implementation.

- **`pptx.Presentation` is a factory function, not the class.** The class is
  `pptx.presentation.Presentation`. Under `mypy --strict` the distinction shows
  up in every signature; import the class under `if TYPE_CHECKING:`.

- **Three different things are called "id."** `sldId/@id` is a deck-scoped
  slide id (≥ 256); the `r:id` on that same element is a **part-scoped**
  relationship id; and the list index is neither. They are never
  interchangeable, and conflating the first two is the second-most common bug
  in this problem space. SPEC §3.3.

- **`r:id` is scoped to the part that contains it.** `rId2` in `slide1.xml` and
  `rId2` in `slide2.xml` are unrelated. This is why `deepcopy` of a slide's XML
  produces dangling references, and it is the reason this library exists.

- **Part names are not identifiers.** `/ppt/slides/slide3.xml` says nothing
  about position — order comes only from `sldIdLst`, and a part named
  `slide3.xml` may not be in the deck at all. Never derive order from a
  filename, and never build a partname with `os.path.join`: OPC part names are
  forward-slashed on every platform.

- **Orphan collection is free; do not write a collector.** `iter_parts()` walks
  the relationship graph and `save` writes only what it reaches, regenerating
  `[Content_Types].xml` from the same tuple. Dropping a relationship *is*
  collecting the part. SPEC §3.5.

- **Assert against the saved package, never the in-memory `Presentation`.**
  The in-memory object graph keeps stale references after a delete by design;
  the serialized package is the artifact that is actually clean. Asserting in
  memory tests the wrong thing.

- **`XmlPart.drop_rel` is conditional and will silently no-op.** It refuses to
  drop a relationship whose reference count in the part's XML is 2 or more —
  and for a slide that also appears in a custom show, it is. Use
  `part.rels.pop(rId)`.

- **`_Relationships.get_or_add` dedupes on `(reltype, target, is_external)`,
  and `_next_rId` fills gaps.** Either one shifts the relationship ids on a
  clone relative to its source, so the id remap is mandatory even when a simple
  case appears to work without it. SPEC §4.4.

- **`SlidePart.notes_slide` and `PresentationPart.notes_master_part` are lazy
  properties that *create* parts** — the latter creates a notes master *and* a
  theme part. Never touch either on an inspection path; use `has_notes_slide`
  and iterate relationships directly. SPEC §9.8.

- **python-pptx does not model `p14:sectionLst` or `p:custShowLst`.** They
  survive round trips only because unrecognized XML is preserved verbatim.
  Generated fixtures never have them, so a change can pass the whole suite and
  still corrupt every deck a real user has organized. `core/sections.py` owns
  this.

## Releasing

```bash
uv run bump-my-version bump {major|minor|patch} --dry-run -v   # preview first; tree must be clean
```

Bumps `pyproject.toml`, `pptx_plus/__init__.py`, and the `uv.lock` entry;
commits and tags `vX.Y.Z`. Pushing the tag runs the full CI gate, verifies the
tag matches `pyproject.toml`, and publishes via PyPI trusted publishing.
`CHANGELOG.md` is maintained by hand.

There is deliberately **no post-release doc re-stamping chore**: no prose file
contains a version literal, so there is nothing to go stale.
