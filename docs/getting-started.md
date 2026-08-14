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
    v0.1 is not released yet. Install from the repository until it is.

```python
from pptx import Presentation
from pptx_plus.slides import delete_slide, duplicate_slide, move_slide

prs = Presentation("deck.pptx")

# Duplicate slide 0. The copy lands immediately after it.
copy = duplicate_slide(prs, 0)

# Move the last slide to the front.
move_slide(prs, -1, 0)

# Remove a slide.
delete_slide(prs, 3)

prs.save("deck.pptx")
```

Every verb accepts either an index or a `Slide` object, and indices follow
Python list semantics — `-1` is the last slide.

For a runnable version with each step verified against the saved package:

```bash
python -m pptx_plus.examples.slide_lifecycle
```

## Asking about a slide

```python
from pptx_plus.slides import contains, slide_index

slide_index(prs, some_slide)   # position, or SlideNotFoundError
contains(prs, some_slide)      # the same question, without the exception
```

A deleted `Slide` object stays alive and readable — deletion detaches its part
from the relationship graph, it destroys nothing — so "is this slide still in
the deck?" is a question you have to ask the deck, not the slide.

## Errors

Every exception subclasses `PptxPlusError` **and** the stdlib type you would
naturally catch:

| Raised | Also a | When |
|---|---|---|
| `SlideNotFoundError` | `KeyError` | The slide is not in this deck |
| `SlideIndexError` | `IndexError` | An index is out of range |
| `DanglingRelationshipError` | `ValueError` | The source XML names a relationship that does not exist |

So `except PptxPlusError` catches everything this library raises, while
existing `except KeyError` clauses keep working. That dual inheritance is also
what makes idempotence opt-in rather than imposed:

```python
import contextlib

with contextlib.suppress(KeyError):
    delete_slide(prs, slide)   # fine if it is already gone
```

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
- [API reference](reference/slides.md) — every public function.
