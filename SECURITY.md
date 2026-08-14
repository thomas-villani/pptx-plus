# Security Policy

## Supported versions

`pptx_plus` is a small library on a rolling release. Security fixes land on the
latest minor version; there are no long-term support branches.

| Version | Supported |
|---|---|
| Latest minor | Yes |
| Anything older | No — upgrade |

## Reporting a vulnerability

**Do not open a public issue for a security problem.**

Report it privately through GitHub's
[private vulnerability reporting](https://github.com/thomas-villani/pptx-plus/security/advisories/new),
or by email to <thomas.villani@gmail.com>.

Please include:

- a description of the issue and its impact,
- a minimal `.pptx` or code sample that reproduces it,
- the `pptx_plus`, `python-pptx`, `lxml`, and Python versions.

You can expect an acknowledgement within a week. This is a small
single-maintainer project, so please size your expectations accordingly — but a
real vulnerability will be prioritised over feature work.

## Threat model

`pptx_plus` reads and rewrites OPC packages, which are ZIP archives. **Treat a
`.pptx` from an untrusted source as untrusted input**, the same way you would
any archive.

Things worth knowing if you process decks you did not author:

- **Decompression bombs.** A `.pptx` is a ZIP, and a hostile one can expand to
  an arbitrary size. `pptx_plus` reads parts through python-pptx, which reads
  through the standard library's `zipfile` — neither imposes a size limit.
  Bound the input yourself if that matters in your environment.
- **XML entity expansion.** XML parts are parsed with `lxml` as configured by
  python-pptx. `pptx_plus` does not relax those settings and does not resolve
  external entities of its own accord.
- **Relationship targets.** Copy operations follow relationships to their
  target parts. External relationships are re-minted by reference and are never
  dereferenced — `pptx_plus` makes no network requests, ever.
- **Path handling.** OPC part names are package-internal and always
  forward-slashed; `pptx_plus` never maps one onto a filesystem path, so a
  crafted part name cannot escape a directory. Files are only written where the
  caller asks python-pptx to save.

A crash or an unhandled exception on a malformed deck is a **bug**, and worth
reporting as one — but it is not by itself a vulnerability. What would be: any
path by which processing a deck causes code execution, network access, or a
write outside the caller's chosen output.
