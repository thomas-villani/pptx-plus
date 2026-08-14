# The OPC model

Everything `pptx_plus` does is a consequence of how a `.pptx` file is put
together. This page is the vocabulary; the rest of the documentation assumes
it.

## A package of parts

A `.pptx` is an **OPC package** — a ZIP archive of **parts**. Each part has a
**part name** (`/ppt/slides/slide1.xml`) and a **content type**, declared
centrally in `[Content_Types].xml`.

```
deck.pptx
├── [Content_Types].xml
├── _rels/.rels
└── ppt/
    ├── presentation.xml
    ├── _rels/presentation.xml.rels
    ├── slides/
    │   ├── slide1.xml
    │   └── _rels/slide1.xml.rels
    ├── media/image1.png
    └── slideLayouts/ …
```

Part names are always forward-slashed, on every platform. They are package
coordinates, not file paths.

## Relationships, and the crux of the whole problem

Parts do not reference each other by name. They reference each other through
**relationships**, stored in a sibling `_rels/<part>.rels` file. Each has an
`Id`, a `Type`, and a `Target`:

```xml
<!-- ppt/slides/_rels/slide1.xml.rels -->
<Relationship Id="rId1" Type=".../slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
<Relationship Id="rId2" Type=".../image"       Target="../media/image1.png"/>
```

The slide's own XML then points at the image by relationship id:

```xml
<a:blip r:embed="rId2"/>
```

Here is the fact everything hinges on:

!!! danger "`r:id` is scoped to the part that contains it"
    `rId2` in `slide1.xml` and `rId2` in `slide2.xml` are **unrelated**. Each
    part has its own relationship namespace, starting again at `rId1`.

## The three things called "id"

| Name | Scope | Where it lives | Range |
|---|---|---|---|
| `p:sldId/@id` | the deck | `presentation.xml` | 256 … 2147483647 |
| `p:sldId/@r:id` | the presentation **part** | the same element | `rIdN` |
| list index | your call site | `prs.slides[i]` | 0 … len−1 |

Three different identifiers, on or near the same element, none
interchangeable. Conflating the first two is the second most common bug in
this problem space — after the one below.

## Order comes from `sldIdLst`, and nowhere else

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId2"/>
  <p:sldId id="257" r:id="rId3"/>
</p:sldIdLst>
```

A part named `slide3.xml` is not necessarily the third slide, and may not be in
the deck at all. Reordering slides means reordering these children — nothing
else moves, and no file is renamed.

Layouts and masters are referenced from `presentation.xml` too, through
`<p:sldLayoutIdLst>` and `<p:sldMasterIdLst>`. That is why deleting the last
slide using a layout does not delete the layout.

## Two side-indexes python-pptx does not model

Both reference slides, and both are invisible to python-pptx — they survive a
round trip only because unrecognized XML is preserved verbatim.

- **`<p:custShowLst>`** — custom shows, referencing slides by **`r:id`**.
- **`<p:extLst>` → `p14:sectionLst`** — PowerPoint sections, referencing slides
  by **`sldId/@id`**.

Leave a stale entry in either and PowerPoint shows a repair prompt on open.
Because generated test decks have neither, this is the one failure mode a full
green test suite can hide — which is why `pptx_plus` maintains both, and why
some of its test fixtures had to be authored in PowerPoint by hand.

## What "orphan" means here

`python-pptx` serializes a package by walking the relationship graph from the
root and writing only the parts it reaches. `[Content_Types].xml` is
regenerated from that same walk.

**So a part that nothing references is simply never written.** Dropping the
presentation part's relationship to a slide *is* the act of collecting it: the
slide part, and any image reachable only through it, do not appear in the saved
file. There is no separate garbage-collection step, and `pptx_plus` deliberately
does not implement one.

One consequence is worth stating on its own, because it shapes how the library
is tested:

!!! warning "The in-memory object is not the artifact"
    After a delete, the in-memory `Presentation` still holds references to the
    removed slide's objects. The **saved package** is what is clean. Any check
    of an operation's effect has to run against a saved-and-reopened deck, or
    it is measuring the wrong thing.

## Why the common recipes are wrong

### The duplicate recipe

```python
new = copy.deepcopy(source.element)          # copies the <p:spTree>
prs.slides._sldIdLst.append(new_sldId)
```

The copied XML still says `r:embed="rId2"`. But `rId2` was scoped to the source
slide's part, and on the new part it means whatever *that* part's relationship
set says — usually nothing. The reference dangles and the picture drops
silently on open.

Fixing it means minting fresh relationships on the new part and **rewriting
every relationship id inside the cloned XML** to match. That rewrite is the
work, and it is what the recipes omit.

There is a trap in verifying it, too. In the simplest case — one image, one
layout — the new relationship ids come out *identical* to the source's, purely
because they were minted in the same order. So an implementation that skips the
rewrite entirely appears to work. It stops working as soon as the source's
relationship ids have a gap in them, which is the ordinary result of having
deleted a shape in PowerPoint.

### The delete recipe

```python
prs.slides._sldIdLst.remove(sld_id)
```

This unlinks the `<p:sldId>` but leaves the presentation part's relationship to
the slide in place. The slide part is therefore still reachable, and is still
written to the saved file. The slide's entries in `p14:sectionLst` and
`p:custShowLst` are left behind as well, dangling.

## So what does `pptx_plus` guarantee?

One invariant, from which the rest follows:

> After any `pptx_plus` operation, every relationship-id-bearing attribute in
> every reachable part resolves to a relationship on that part.

It is checked after every operation in the test suite, on the saved package,
including against decks PowerPoint authored with SmartArt, video, sections, and
custom shows.
