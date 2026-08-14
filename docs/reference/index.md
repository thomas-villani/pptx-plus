# API reference

Two layers, and the boundary between them is meaningful.

**[`pptx_plus.slides`](slides.md)** is the public API: three verbs and the
helpers that resolve their arguments. If you are using this library, this is
the page you want.

**[`pptx_plus.core`](core.md)** is the foundation the verbs are built from —
the copy engine, the relationship-id rewriter, part allocation, and the
maintenance of the side-indexes python-pptx does not model. It is documented
because the reasoning in it is the substance of the library, and because
anyone extending `pptx_plus` works here. It is *not* covered by the same
stability promise as `slides`.

**[`pptx_plus._testing`](testing.md)** ships the OPC integrity battery. It is
test-only and unstable, but installed rather than left in `tests/`, so a
project that wraps this library can assert the same invariants against its own
decks.

Each subpackage's `__init__.py` `__all__` is the authoritative surface for that
package.
