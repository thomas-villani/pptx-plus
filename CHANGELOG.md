# Changelog

All notable changes to `pptx_plus` are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Entries explain *why*, not only *what*. A one-line "fixed a bug" tells a
reader nothing they can act on.

## [Unreleased]

### Added

- Repository scaffolding: packaging, tooling, CI, the specification, and the
  implementation notes. No capability code yet — see `ROADMAP.md` for the v0.1
  target.
- `pptx_plus.core`: the foundation layer — namespaces, the element chokepoint,
  relationship-type constants and clone policy sets, slide-id allocation,
  part-name allocation and byte-faithful part cloning, and an upstream-surface
  guard that fails at import with the name of any python-pptx attribute this
  library depends on that has moved.
- `pptx_plus.slides.move_slide` — reorder a slide, with `to_index` naming its
  position in the *resulting* deck. Out of range raises rather than clamping:
  `list.insert` clamps, which would turn an off-by-one in a caller's loop into
  a silently misordered deck.
- `pptx_plus.slides`: `resolve_slide`, `slide_index`, and `contains`, plus the
  `SlideNotFoundError` / `SlideIndexError` pair. Both dual-inherit the stdlib
  exception a caller would naturally catch, so `except KeyError` still works
  and `contextlib.suppress(KeyError)` gives opt-in idempotence.
- `pptx_plus.core.sections` — maintenance for `p14:sectionLst` and
  `p:custShowLst`, neither of which python-pptx models. They survive its round
  trips only because unrecognized XML is preserved verbatim, so an operation
  that edits the slide list and stops leaves them naming a slide that is gone
  and PowerPoint reports the file as damaged.
- `pptx_plus._testing`: the OPC integrity battery (SPEC §10.3), which reads a
  saved package as a zip rather than through python-pptx — the in-memory object
  graph keeps stale references after a delete by design, so asserting against
  it grades the wrong artifact. Shipped in the wheel so downstream projects can
  assert the same invariants against their own decks.
