"""Prose documentation must not carry a version literal of its own.

docx-plus, this project's sibling, carries a documented post-release chore to
re-stamp the version in six prose files — and records that the chore "has
historically lagged behind the bump." A documented chore that reliably lags is
a design bug, not a discipline problem, so the fix here is structural: prose
simply does not state the version.

Version literals are confined to the five places the release process already
touches — `pyproject.toml`, `pptx_plus/__init__.py`, `uv.lock`,
`CHANGELOG.md` headings, and ROADMAP's current-state line. The README's PyPI
badge renders the live version, which is strictly better than a hand-typed one.

SPEC §12.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent

# A version literal: three dot-separated numbers, optionally `v`-prefixed, not
# part of a longer dotted run.
#
# The two lookaheads are doing different jobs and both are needed:
#   (?!\.?\d)  keeps ECMA-376 clause numbers like §19.2.1.34 from matching on
#              their first three segments — but still permits a trailing
#              sentence period, so "released as v1.2.3." is caught.
#   (?!\w)     keeps 1.2.3rc1 and the like from matching on the numeric head.
_VERSION = re.compile(r"(?<![\w.])v?\d+\.\d+\.\d+(?!\.?\d)(?!\w)")

# Fenced code blocks are exempt: an install command or a sample `.rels` file
# may legitimately contain a version, and neither goes stale in a way a reader
# would be misled by.
_FENCE = re.compile(r"^\s*(```|~~~)")

# Version literals that are *about something else* — a dependency floor, an
# external standard, a third-party tool. These describe another project's
# version, so they do not go stale when this project releases.
_ALLOWED_CONTEXT = (
    "python-pptx",
    "python-docx",
    "lxml",
    "Python 3",
    "ECMA-376",
    "Keep a Changelog",
    "keepachangelog",
    "Contributor Covenant",
    "semver.org",
    "LibreOffice",
    "PowerPoint",
    "Word",
)


def _prose_files() -> list[Path]:
    files = [REPO_ROOT / "README.md", REPO_ROOT / "SPEC.md"]
    files.extend(sorted((REPO_ROOT / "docs").rglob("*.md")))
    return [p for p in files if p.is_file()]


def _offending_lines(path: Path) -> list[tuple[int, str]]:
    lines = path.read_text(encoding="utf-8").splitlines()
    offenders: list[tuple[int, str]] = []
    in_fence = False
    for index, line in enumerate(lines):
        if _FENCE.match(line):
            in_fence = not in_fence
            continue
        if in_fence or not _VERSION.search(line):
            continue
        # Prose is hard-wrapped, so the word that gives a number its meaning
        # ("ECMA-376", "python-pptx") often sits on the line above or below the
        # number itself. Match the allowlist against the enclosing paragraph —
        # the run of non-blank lines around this one — rather than the line.
        start = index
        while start > 0 and lines[start - 1].strip():
            start -= 1
        end = index
        while end + 1 < len(lines) and lines[end + 1].strip():
            end += 1
        paragraph = "\n".join(lines[start : end + 1])
        if any(token in paragraph for token in _ALLOWED_CONTEXT):
            continue
        offenders.append((index + 1, line.strip()))
    return offenders


@pytest.mark.parametrize("path", _prose_files(), ids=lambda p: str(p.relative_to(REPO_ROOT)))
def test_prose_carries_no_version_literal(path: Path) -> None:
    offenders = _offending_lines(path)
    assert not offenders, (
        f"{path.relative_to(REPO_ROOT)} contains a version literal — prose must not "
        f"state this project's version (SPEC §12). Offending lines: {offenders}"
    )


def test_the_check_can_actually_fail(tmp_path: Path) -> None:
    """Guard the guard.

    A regex this fussy — a negative lookahead, a fence toggle, and a context
    allowlist — can easily end up matching nothing at all and passing every
    file vacuously. This asserts it still catches the thing it is for.
    """
    doc = tmp_path / "sample.md"
    doc.write_text(
        "# Title\n"
        "\n"
        "```\n"
        "pip install pptx-plus==9.9.9\n"  # inside a fence: exempt
        "```\n"
        "\n"
        "See ECMA-376 for the clause numbered\n"  # the context word...
        "§19.3.1, which looks exactly like a version.\n"  # ...wraps onto here
        "\n"
        "Requires python-pptx 1.0.2 or newer.\n"  # allowlisted dependency floor
        "\n"
        "Current release is v9.9.9.\n",  # the one real offender
        encoding="utf-8",
    )
    assert _offending_lines(doc) == [(12, "Current release is v9.9.9.")]
