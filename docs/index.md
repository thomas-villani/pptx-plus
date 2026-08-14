# pptx_plus

**OPC-level extensions for [python-pptx](https://python-pptx.readthedocs.io/).**

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
inside it pointing at the **source slide's** relationships. Relationship ids
are scoped to the part that contains them, so on the new part they resolve to
something else, or to nothing. Pictures, charts, and media are
relationship-targeted parts, so the copy carries dangling references and the
picture silently vanishes when the deck is opened.

Relationship and part rewriting is the actual work, and every shortcut skips
it.

```python
from pptx import Presentation
from pptx_plus.slides import duplicate_slide

prs = Presentation("deck.pptx")
copy = duplicate_slide(prs, 0, to_index=1)
prs.save("deck.pptx")
```

## How it works

`pptx_plus` operates at the **OPC layer** — the package, its parts, and the
relationships between them — below python-pptx's high-level object model.

That choice has a useful consequence. Because a duplicated slide's parts are
cloned byte-for-byte rather than reconstructed through the object model,
content python-pptx does not understand — SmartArt, embedded video, OLE
objects, custom XML — comes along intact. It never passes through the layer
that would strip it.

That is preservation, not authoring. `pptx_plus` does not create SmartArt, and
does not yet read or edit it.

## Where to go next

- **[Getting started](getting-started.md)** — install, and your first
  duplicate.
- **[Concepts: the OPC model](concepts/opc-model.md)** — parts, relationships,
  the three things called "id," and precisely why the common recipes fail.
- **[Roadmap](https://github.com/thomas-villani/pptx-plus/blob/main/ROADMAP.md)**
  — what's coming, and what is refused.

Sister project: **[docx-plus](https://github.com/thomas-villani/docx-plus)**,
the same idea for `python-docx`.
