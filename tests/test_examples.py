"""The shipped examples actually run -- SPEC §11."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

EXAMPLES = ["pptx_plus.examples.slide_lifecycle"]


@pytest.mark.parametrize("module", EXAMPLES)
def test_the_example_runs(module: str, tmp_path: Path) -> None:
    result = subprocess.run(
        [sys.executable, "-m", module, "--out", str(tmp_path)],
        capture_output=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr.decode(errors="replace")


@pytest.mark.parametrize("module", EXAMPLES)
def test_the_example_writes_a_deck(module: str, tmp_path: Path) -> None:
    subprocess.run(
        [sys.executable, "-m", module, "--out", str(tmp_path)],
        capture_output=True,
        check=True,
    )
    assert list(tmp_path.glob("*.pptx"))


@pytest.mark.parametrize("module", EXAMPLES)
def test_the_example_output_is_cp1252_safe(module: str, tmp_path: Path) -> None:
    """Runs on a default Windows console, whose code page is cp1252.

    A typographic quote or an arrow in a print raises `UnicodeEncodeError` at
    the print, nowhere near the code that produced the string -- so this is
    checked by actually encoding the output rather than by reading the source.
    """
    result = subprocess.run(
        [sys.executable, "-m", module, "--out", str(tmp_path)],
        capture_output=True,
        check=True,
    )
    result.stdout.decode("utf-8").encode("cp1252")


@pytest.mark.parametrize("module", EXAMPLES)
def test_the_example_writes_nowhere_it_was_not_asked_to(
    module: str, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Run from an empty cwd with --out elsewhere: nothing lands in cwd."""
    workdir = tmp_path / "cwd"
    workdir.mkdir()
    out = tmp_path / "out"
    monkeypatch.chdir(workdir)
    subprocess.run(
        [sys.executable, "-m", module, "--out", str(out)],
        capture_output=True,
        check=True,
        cwd=workdir,
    )
    assert list(workdir.iterdir()) == []
