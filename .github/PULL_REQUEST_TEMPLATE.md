<!--
Thanks for contributing. Keep the change focused on one concern; see
CONTRIBUTING.md for the conventions and quality gates.
-->

## What this changes

<!-- A sentence or two. If it closes an issue, say "Closes #123". -->

## Why

<!--
The reasoning, not just the mechanics. If this fills a python-pptx gap, name
it. If it changes existing OOXML output, say what PowerPoint did before and
what it does now.
-->

## How it was verified

<!--
Delete what does not apply.
-->

- [ ] New tests cover the change
- [ ] **The relationship graph closes** — `assert_rel_ids_resolve` and
      `assert_no_unclaimed_rid_literals` pass on the saved package
- [ ] Assertions run against a **saved and reopened** deck, not the in-memory
      `Presentation` (the in-memory graph keeps stale references by design)
- [ ] Verified the OOXML against a file **PowerPoint itself authored** (see
      CONTRIBUTING.md § Verifying against PowerPoint)
- [ ] Opened the output in PowerPoint and confirmed **no repair prompt**
- [ ] Opened the output in LibreOffice / converted it headlessly

<!-- If you could not verify against PowerPoint, say so here — flagging the
     gap is better than asserting an unverified shape. -->

## Checklist

- [ ] `uv run pytest` passes
- [ ] `uv run mypy` passes (strict, zero ignores)
- [ ] `uv run ruff check` and `uv run ruff format` are clean
- [ ] `uv run mkdocs build --strict` builds link-clean
- [ ] `CHANGELOG.md` updated under `## [Unreleased]`
- [ ] Public symbols are exported in the subpackage `__all__` and listed in
      `docs/API.md`
- [ ] `ROADMAP.md` updated, if this lands or reshapes a roadmap item
- [ ] No version literal added to `README.md`, `SPEC.md`, or `docs/**` (see
      `tests/test_no_version_literals.py`)

## Breaking changes

<!-- None, or describe the migration. -->
