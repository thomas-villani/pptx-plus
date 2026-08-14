"""Fixture decks, built once per session into a temp directory.

Decks are generated rather than committed, and generated *outside* the source
tree -- `tmp_path_factory` puts them under pytest's own base temp directory.
A test that mutates a deck opens it from bytes, so the built files stay
pristine and the session scope is safe.
"""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

import pytest

from tests.fixtures.build_decks import BUILDERS

if TYPE_CHECKING:
    from collections.abc import Callable


@pytest.fixture(scope="session")
def fixture_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """The directory generated fixture decks are built into."""
    return tmp_path_factory.mktemp("decks")


@pytest.fixture(scope="session")
def deck(fixture_dir: Path) -> Callable[[str], Path]:
    """Return a builder that produces a named fixture deck on first request.

    Memoized on the path: a deck is built once per session however many tests
    ask for it, but a deck no test asks for is never built at all.

        def test_something(deck):
            path = deck("picture")
    """

    def build(name: str) -> Path:
        try:
            builder = BUILDERS[name]
        except KeyError:
            raise AssertionError(
                f"unknown fixture deck {name!r}; registered: {sorted(BUILDERS)}"
            ) from None
        path = fixture_dir / f"{name}.pptx"
        if not path.exists():
            builder(path)
        return path

    return build


@pytest.fixture(scope="session")
def sample_dir() -> Path:
    """The directory of committed, PowerPoint-authored sample decks.

    Unlike the generated fixtures these are binaries in the repository,
    because SmartArt, embedded video, sections, and custom shows cannot be
    authored by python-pptx at all. Provenance is recorded in the README
    alongside them. SPEC §10.6.
    """
    return Path(__file__).parent / "fixtures" / "pptx_samples"
