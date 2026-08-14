# Contributing to pptx_plus

Thanks for your interest. This is a small, single-maintainer project with
opinionated conventions; reading this first will save you a review round.

## Scope

`pptx_plus` is a **lean extension to python-pptx**, not a deck-authoring
framework. Every capability here either fills a documented python-pptx gap or
rounds out a surface already started.

Before proposing a feature, check `ROADMAP.md` — including its *Considered, not
on the roadmap* section, which refuses animations, a SmartArt layout engine,
templating, and a `Presentation` subclass explicitly rather than by omission.

`SPEC.md` is the API contract. A change that alters documented behaviour has to
change SPEC in the same PR.

## Development setup

This project uses **`uv`** for everything. Never call bare `python` or `pip`.

```bash
git clone https://github.com/thomas-villani/pptx-plus.git
cd pptx-plus
uv sync --extra dev
uv run pre-commit install
uv run pytest
```

## Quality gates

All of these must pass before a PR is mergeable. CI runs them on Linux
3.10–3.13 and on Windows 3.13, plus a lower-bound dependency leg.

```bash
uv run ruff check pptx_plus/ tests/
uv run ruff format --check pptx_plus/ tests/
uv run mypy --strict pptx_plus/
uv run pytest --cov=pptx_plus --cov-fail-under=90
uv run mkdocs build --strict
```

`mypy` is strict with zero ignores. Coverage is enforced at 90%.

## Conventions

- **`from __future__ import annotations`** at the top of every module;
  `if TYPE_CHECKING:` for python-pptx types.
- **Keyword-only arguments after `*`** on public functions.
- **Google-convention docstrings.** A module docstring names the python-pptx
  gap it fills, cites the relevant ECMA-376 clause, and ends with the invariant
  line ``This module imports only from ``pptx_plus.core`` (SPEC §9.1).``
- **Absolute imports only.** No relative imports anywhere — the AST check in
  `tests/test_import_invariant.py` matches on absolute module names and
  relative imports would defeat it.
- **No cross-capability imports.** Capability packages depend on `core/` and on
  nothing else in this library.
- **All element construction goes through `core/oxml.py`.** No bare
  `lxml.etree` calls in a capability module.
- **Typed errors** subclass `PptxPlusError` and dual-inherit the stdlib type a
  caller would naturally catch. See `SPEC.md §16`.
- **No version literals** in `README.md`, `SPEC.md`, or under `docs/`.
- **Comments explain why.** A comment restating the code is noise; a comment
  recording why the obvious approach is wrong is the most valuable thing in the
  file. Much of this codebase exists because the obvious approach is wrong.

## Verifying against PowerPoint

For anything touching a part, a relationship, or an attribute this project has
not written before, the expectation is to **have PowerPoint author the file
first** — unzip it, and match the observed markup. Guessing at OOXML shapes
produces code that validates and renders wrong.

`tests/fixtures/pptx_samples/README.md` documents how each committed sample
deck was produced. If you add one, add its provenance there: PowerPoint
version, OS, date, and the exact click-path.

## And verify the relationship graph closes

This is the half with no docx-plus analogue, and it is the point of the
project. Every operation must leave every relationship-id-bearing attribute in
every reachable part resolving to a relationship on that part.

Two rules that follow, and that reviewers will ask about:

1. **Assert against a saved and reopened package**, never the in-memory
   `Presentation`. The in-memory object graph keeps stale references by design;
   the serialized package is the artifact users get. Use
   `pptx_plus._testing.roundtrip`.
2. **Call the battery.** `assert_rel_ids_resolve` and
   `assert_no_unclaimed_rid_literals` catch the failure mode every
   StackOverflow slide-copy recipe has, and they are cheap. See `SPEC.md §10.3`.

A test that passes on the first attempt deserves a moment's suspicion here.
The failure modes in this problem space are silent, and at least one of them —
the relationship-id remap — produces an *identity* mapping in the simple case,
so an implementation that skips it entirely passes the obvious test. Make a
test fail on purpose before trusting that it passes.

## Pull requests

- One concern per PR.
- Conventional commit subjects (`feat:`, `fix:`, `docs:`, `test:`, `chore:`).
- Update `CHANGELOG.md` under `## [Unreleased]`, explaining *why*, not only
  what.
- Update `ROADMAP.md` if the change lands or reshapes a roadmap item.
- Export new public symbols in the subpackage `__all__` and list them in
  `docs/API.md`.

## Reporting bugs

The most useful attachment is a **minimal `.pptx`** that reproduces the
problem. If the deck is sensitive, the next best thing is the shape of the
package rather than its content: `unzip -l deck.pptx` plus the offending
`ppt/slides/_rels/slideN.xml.rels`. Most bugs here are relationship-graph bugs,
and those are fully visible in the rels file without exposing any of your text.

Please also say whether the deck uses sections, custom shows, SmartArt,
embedded media, or slide-jump hyperlinks — those are the features python-pptx
does not model, and they are disproportionately where things break.

## Releasing

Maintainer-only, recorded here so the process is not folklore.

```bash
uv run bump-my-version bump {major|minor|patch} --dry-run -v   # preview; tree must be clean
uv run bump-my-version bump minor
git push --follow-tags
```

The tag push runs the full CI gate, verifies the tag matches
`pyproject.toml`, and publishes via PyPI trusted publishing. `CHANGELOG.md` is
maintained by hand.

There is no post-release documentation chore: no prose file carries a version
literal, so nothing goes stale.

## License

By contributing you agree that your contributions are licensed under the MIT
License, the same as the project.
