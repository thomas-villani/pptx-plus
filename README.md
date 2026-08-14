<div align="center">

# pptx_plus

**OPC-level extensions for [python-pptx](https://python-pptx.readthedocs.io/).**

[![PyPI](https://img.shields.io/pypi/v/pptx-plus.svg?logo=pypi&logoColor=white)](https://pypi.org/project/pptx-plus/)
[![Python versions](https://img.shields.io/pypi/pyversions/pptx-plus.svg?logo=python&logoColor=white)](https://pypi.org/project/pptx-plus/)
[![CI](https://github.com/thomas-villani/pptx-plus/actions/workflows/ci.yml/badge.svg)](https://github.com/thomas-villani/pptx-plus/actions/workflows/ci.yml)
[![Docs](https://img.shields.io/badge/docs-mkdocs--material-blue)](https://thomas-villani.github.io/pptx-plus/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/thomas-villani/pptx-plus/blob/main/LICENSE)
[![Typed](https://img.shields.io/badge/typing-strict-blue)](https://mypy-lang.org/)
[![Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)

[Documentation](https://thomas-villani.github.io/pptx-plus/) ·
[Getting started](https://thomas-villani.github.io/pptx-plus/getting-started/) ·
[Guides](https://thomas-villani.github.io/pptx-plus/guides/) ·
[API index](https://thomas-villani.github.io/pptx-plus/API/) ·
[Changelog](https://github.com/thomas-villani/pptx-plus/blob/main/CHANGELOG.md) ·
[Roadmap](https://github.com/thomas-villani/pptx-plus/blob/main/ROADMAP.md)

</div>

---

python-pptx is an excellent library with one conspicuous hole: **it cannot
delete, move, or duplicate a slide.** Those are its two oldest open feature
requests — [#67](https://github.com/scanny/python-pptx/issues/67) and
[#132](https://github.com/scanny/python-pptx/issues/132) — and they have been
open for over a decade.

The recipes that fill the gap are not merely unofficial; they are wrong.

```python
# The duplicate recipe everyone copies from StackOverflow.
new = copy.deepcopy(source.element)
prs.slides._sldIdLst.append(new_sldId)
```

That clones the slide's `<p:spTree>` and leaves every `r:id` and `r:embed`
inside it pointing at the **source slide's** relationships. Pictures, charts,
and media are relationship-targeted parts, so the copy carries dangling
references and the picture silently vanishes when the deck is opened.

Relationship and part rewriting is the actual work, and every shortcut skips
it. `pptx_plus` does it properly, at the OPC layer — below python-pptx's
object model, so nothing round-trips through a lossy re-serialization on the
way.

```python
from pptx import Presentation
from pptx_plus.slides import duplicate_slide

prs = Presentation("deck.pptx")

# Every r:id in the cloned XML is rewritten to the clone's own relationships.
# The image is shared by reference; the chart and its embedded workbook are
# deep-cloned; the notes come along.
copy = duplicate_slide(prs, 0, to_index=1)

prs.save("deck.pptx")
```

## Install

```bash
pip install pptx-plus
```

## What it does

`pptx_plus` composes with python-pptx rather than replacing it. You keep your
`Presentation` object and reach for `pptx_plus` only where python-pptx stops.

| Verb | What it handles that the recipes don't |
|---|---|
| `delete_slide` | Drops the relationship *and* scrubs the slide from PowerPoint sections and custom shows — the two side-indexes python-pptx does not model, and the usual cause of a "PowerPoint found a problem" repair prompt. |
| `move_slide` | Reorders `<p:sldIdLst>` and relocates the slide's entry in any section. Pure ordering; no part changes. |
| `duplicate_slide` | Clones the slide's whole part graph — images shared by reference, charts and their embedded workbooks deep-cloned — and rewrites every relationship id inside the cloned XML. |

### A fidelity side effect worth knowing about

Because the copy engine works on **parts and relationships** rather than
reconstructing content through python-pptx's object model, anything on a
duplicated slide that python-pptx doesn't understand — SmartArt, embedded
video, OLE objects, custom XML — is cloned byte-for-byte and **survives**. It
never passes through the model that would strip it.

That is preservation, not authoring: `pptx_plus` does not create SmartArt, and
v0.1 does not read or edit it either. See
[`ROADMAP.md`](https://github.com/thomas-villani/pptx-plus/blob/main/ROADMAP.md).

## Not in scope

Stated up front, because these are what people ask for first:

- **Animations and transitions.** An XML rat's nest with no good abstraction.
- **SmartArt authoring.** The layout part is a declarative layout *program*
  that PowerPoint executes to compute geometry. Reimplementing that engine is
  a permanent non-goal; a template-clone authoring path is on the roadmap.
- **A `Presentation` subclass**, or anything that replaces python-pptx's model.

## For AI coding agents

`pptx_plus` is built to back LLM tooling that generates and manipulates decks,
which shapes two things about its design: the public surface is small and
verb-shaped, and every operation is verified against the saved package rather
than the in-memory object graph, so an agent cannot be told an operation
succeeded when the resulting file is corrupt.

## Documentation

Full documentation is at
[thomas-villani.github.io/pptx-plus](https://thomas-villani.github.io/pptx-plus/).

- **[Getting started](https://thomas-villani.github.io/pptx-plus/getting-started/)** — install and first duplicate.
- **[Guides](https://thomas-villani.github.io/pptx-plus/guides/)** — task-oriented.
- **[Concepts](https://thomas-villani.github.io/pptx-plus/concepts/)** — the OPC model, why the recipes are wrong, and what the library guarantees.
- **[API index](https://thomas-villani.github.io/pptx-plus/API/)** — every public symbol.

## Project status

Alpha. The public surface is small on purpose: v0.1 is one coherent capability
— the slide lifecycle — done correctly, rather than a broad but shallow
feature list.

Cross-deck `copy_slide` and `merge_presentations` are the v0.2 targets; see
[`ROADMAP.md`](https://github.com/thomas-villani/pptx-plus/blob/main/ROADMAP.md)
for the live picture and
[`SPEC.md`](https://github.com/thomas-villani/pptx-plus/blob/main/SPEC.md) for
the API contract.

Sister project: **[docx-plus](https://github.com/thomas-villani/docx-plus)**,
the same idea for `python-docx`.

## Contributing

Contributions are welcome — see
[`CONTRIBUTING.md`](https://github.com/thomas-villani/pptx-plus/blob/main/CONTRIBUTING.md)
for the development setup, the quality gates, and the conventions. In short:

```bash
git clone https://github.com/thomas-villani/pptx-plus.git
cd pptx-plus
uv sync --extra dev
uv run pre-commit install
uv run pytest
```

Bug reports are most useful with a **minimal `.pptx`** attached — or, if the
deck is sensitive, the output of `unzip -l` on it plus the offending
`ppt/slides/_rels/slideN.xml.rels`.

Security issues should be reported privately — see
[`SECURITY.md`](https://github.com/thomas-villani/pptx-plus/blob/main/SECURITY.md).

## License

MIT. Copyright (c) 2026 Tom Villani, PhD. See [`LICENSE`](https://github.com/thomas-villani/pptx-plus/blob/main/LICENSE).
