# Getting started

## Install

```bash
pip install pptx-plus
```

Or, with `uv`:

```bash
uv add pptx-plus
```

python-pptx comes along as a dependency. `pptx_plus` **composes with** it
rather than replacing it — you keep your `Presentation` object and reach for
`pptx_plus` only where python-pptx stops.

## The three verbs

!!! note "In development"
    v0.1 is not released yet. The API below is the contract from
    [`SPEC.md`](https://github.com/thomas-villani/pptx-plus/blob/main/SPEC.md);
    this page will gain runnable output once the verbs land.

```python
from pptx import Presentation
from pptx_plus.slides import delete_slide, duplicate_slide, move_slide

prs = Presentation("deck.pptx")

# Duplicate slide 0 and place the copy immediately after it.
copy = duplicate_slide(prs, 0, to_index=1)

# Move the last slide to the front.
move_slide(prs, -1, 0)

# Remove a slide.
delete_slide(prs, 3)

prs.save("deck.pptx")
```

Every verb accepts either an index or a `Slide` object, and indices follow
Python list semantics — `-1` is the last slide.

## Two things worth knowing up front

### `to_index` is a position in the *resulting* deck

`to_index` behaves like `list.insert` against the list with the slide already
removed. So `move_slide(prs, i, i)` is a no-op for every `i`:

| Deck | Call | Result |
|---|---|---|
| A B C D | `move_slide(prs, 0, 2)` | B C **A** D |
| A B C D | `move_slide(prs, 2, 2)` | A B C D |
| A B C D | `move_slide(prs, 3, 0)` | **D** A B C |

Unlike `list.insert`, an out-of-range `to_index` raises rather than clamping.
Clamping hides off-by-one bugs in caller loops, which is exactly the bug class
this library exists to eliminate.

### What a duplicate shares, and what it copies

| Content | Treatment |
|---|---|
| Images, video, audio, fonts | **Shared** — the copy points at the same part, as PowerPoint's own duplicate does |
| Charts (and their embedded workbooks) | **Deep-cloned** — editing the copy's chart does not touch the original's |
| SmartArt, OLE objects, custom XML | **Deep-cloned**, byte-for-byte |
| Speaker notes | **Deep-cloned**, unless you pass `with_notes=False` |

Every relationship id inside the cloned XML is rewritten to point at the
copy's own relationships. That rewrite is the part every hand-rolled recipe
omits, and the reason the copy's pictures survive.

## Next

- [Concepts: the OPC model](concepts/opc-model.md) — what a `.pptx` actually
  is, and why relationship ids are the crux.
