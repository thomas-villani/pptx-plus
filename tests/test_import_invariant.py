"""SPEC §9.1 — the layering invariant, enforced rather than documented.

Two rules, checked by walking each module's AST:

1. **No imports between capability packages.** A capability depends on
   `core/` and on nothing else in this library. Left to convention this decays
   the first time two capabilities have something almost in common; the right
   answer then is to move the shared thing into `core/`, and this test is what
   forces that conversation to happen.

2. **No relative imports, anywhere.** Not a style preference: rule 1 is
   checked by matching absolute module names, and `from ..core import qn`
   carries no absolute name to match. One relative import would create a hole
   in the check big enough to drive rule 1 through.

`CAPABILITIES` has one member today, which looks like ceremony. It costs
nothing, and it is what stops `slides/` from importing `sections/` in a later
cycle — by which point the convention would exist only in someone's memory.

Composing layers — a future `cli/` — would be excluded here, since composing
across capabilities is exactly their job.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

PACKAGE_ROOT = Path(__file__).resolve().parent.parent / "pptx_plus"

#: Capability packages. `core` is the foundation and `_testing` is test-only;
#: neither is a capability, and neither is subject to rule 1.
CAPABILITIES = frozenset({"slides"})

#: Layers that legitimately import across capabilities. Empty at v0.1.
COMPOSING_LAYERS: frozenset[str] = frozenset()


def _python_files() -> list[Path]:
    return sorted(
        path
        for path in PACKAGE_ROOT.rglob("*.py")
        if "__pycache__" not in path.parts and path.name != "__init__.py"
    )


def _capability_of(path: Path) -> str | None:
    relative = path.relative_to(PACKAGE_ROOT)
    top = relative.parts[0]
    return top if top in CAPABILITIES else None


def _imported_modules(tree: ast.AST) -> list[str]:
    names: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names.extend(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
            names.append(node.module)
    return names


def _relative_imports(tree: ast.AST) -> list[str]:
    return [
        "." * node.level + (node.module or "")
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom) and node.level > 0
    ]


_FILES = _python_files()
_IDS = [str(p.relative_to(PACKAGE_ROOT)).replace("\\", "/") for p in _FILES]


@pytest.mark.parametrize("path", _FILES, ids=_IDS)
def test_no_relative_imports(path: Path) -> None:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    assert not _relative_imports(tree), (
        f"{path.name} uses a relative import. Absolute imports only — the "
        f"cross-capability check below matches on absolute module names and "
        f"cannot see a relative one. SPEC §9.1."
    )


@pytest.mark.parametrize("path", _FILES, ids=_IDS)
def test_no_cross_capability_imports(path: Path) -> None:
    capability = _capability_of(path)
    if capability is None or capability in COMPOSING_LAYERS:
        pytest.skip("not a capability module")

    tree = ast.parse(path.read_text(encoding="utf-8"))
    offenders = [
        name
        for name in _imported_modules(tree)
        for other in CAPABILITIES - {capability}
        if name == f"pptx_plus.{other}" or name.startswith(f"pptx_plus.{other}.")
    ]
    assert not offenders, (
        f"{path.name} (capability {capability!r}) imports another capability: "
        f"{offenders}. Shared code belongs in pptx_plus.core. SPEC §9.1."
    )


def test_core_does_not_import_a_capability() -> None:
    """The dependency arrow points one way only.

    Checked as a whole rather than per-file so the failure names every
    offender at once — a cycle introduced here is usually more than one edge.
    """
    offenders: list[str] = []
    for path in _FILES:
        if path.relative_to(PACKAGE_ROOT).parts[0] != "core":
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"))
        offenders.extend(
            f"{path.name} -> {name}"
            for name in _imported_modules(tree)
            for capability in CAPABILITIES
            if name.startswith(f"pptx_plus.{capability}")
        )
    assert not offenders, f"core/ must not depend on a capability: {offenders}"


def test_the_walk_actually_finds_files() -> None:
    """Guard the guard.

    Every check above is parametrized over `_python_files()`. If that ever
    returns nothing — a renamed package directory, a path bug on one platform
    — the whole file would pass green while checking precisely nothing.
    """
    assert len(_FILES) >= 3
