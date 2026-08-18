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


def test_every_package_function_parameter_is_documented() -> None:
    """Prevent package function and method parameters from being undocumented."""

    missing: list[str] = []
    for path in sorted(PACKAGE.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, (ast.AsyncFunctionDef, ast.FunctionDef)):
                continue
            parameters = [
                *node.args.posonlyargs,
                *node.args.args,
                *node.args.kwonlyargs,
            ]
            if node.args.vararg is not None:
                parameters.append(node.args.vararg)
            if node.args.kwarg is not None:
                parameters.append(node.args.kwarg)
            docstring = ast.get_docstring(node) or ""
            for parameter in parameters:
                if parameter.arg in {"self", "cls"}:
                    continue
                if f"{parameter.arg}:" not in docstring:
                    missing.append(
                        f"{path.relative_to(ROOT)}:{node.lineno}:{node.name}({parameter.arg})"
                    )

    assert not missing, "missing parameter descriptions:\n" + "\n".join(missing)
