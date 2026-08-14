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


def pytest_configure(config: pytest.Config) -> None:
    """Promote `RelIdLiteralWarning` to an error for the whole suite.

    `RelIdLiteralWarning` is the tripwire for a relationship-id attribute the
    rewriter does not know about (SPEC §4.4). At runtime it stays a warning,
    because a shape genuinely named "rId7" must not crash a library call. In
    the test suite it has to be fatal -- that is how the next
    `dsp:dataModelExt/@relId` gets found instead of silently producing a
    dangling reference in someone's deck.

    Registered here rather than as a `filterwarnings` entry in
    `pyproject.toml` because pytest resolves an ini filter by importing the
    named class at config time. That imports `pptx_plus` before coverage
    starts, so every module-level statement in the package records as
    unexecuted and the total reads ~60% instead of ~98% -- a coverage gate
    that fails for a reason having nothing to do with the tests. Appending the
    filter from `pytest_configure` leaves the import to the warnings plugin,
    which parses filters per test item, long after coverage is running.
    """
    config.addinivalue_line("filterwarnings", "error::pptx_plus.core.relmap.RelIdLiteralWarning")


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
