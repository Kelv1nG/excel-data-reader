from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).parents[1]
PACKAGE = ROOT / "src" / "excel_data_reader"


def test_every_package_function_has_a_docstring() -> None:
    """Prevent new package functions or methods from being left undocumented."""

    missing: list[str] = []
    for path in sorted(PACKAGE.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if (
                isinstance(node, (ast.AsyncFunctionDef, ast.FunctionDef))
                and ast.get_docstring(node) is None
            ):
                missing.append(f"{path.relative_to(ROOT)}:{node.lineno}:{node.name}")

    assert not missing, "missing function docstrings:\n" + "\n".join(missing)
