"""Static artifact dependency discovery for recovered Python generators."""

from __future__ import annotations

import ast
from dataclasses import dataclass, field
from pathlib import PurePosixPath


READ_METHODS = {"read_bytes", "read_text", "open"}
WRITE_METHODS = {"write_bytes", "write_text", "save", "savefig", "imwrite"}
READ_FUNCTIONS = {"open", "Image.open", "cv2.imread", "pandas.read_csv", "pandas.read_excel", "pd.read_csv", "pd.read_excel"}


def _literal_path(node: ast.AST) -> str | None:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "Path" and node.args:
        return _literal_path(node.args[0])
    return None


def _call_name(node: ast.AST) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        prefix = _call_name(node.value)
        return f"{prefix}.{node.attr}" if prefix else node.attr
    return ""


def _normalise(value: str) -> str:
    return PurePosixPath(value.replace("\\", "/")).as_posix()


@dataclass(slots=True)
class PythonArtifactSpec:
    inputs: set[str] = field(default_factory=set)
    outputs: set[str] = field(default_factory=set)


class _Visitor(ast.NodeVisitor):
    def __init__(self) -> None:
        self.spec = PythonArtifactSpec()

    def visit_Call(self, node: ast.Call) -> None:
        name = _call_name(node.func)
        method = name.rsplit(".", 1)[-1]

        if isinstance(node.func, ast.Attribute):
            owner_path = _literal_path(node.func.value)
            if owner_path and method in READ_METHODS | WRITE_METHODS:
                mode = ""
                if method == "open":
                    if len(node.args) > 0:
                        mode = str(_literal_path(node.args[0]) or "")
                    for keyword in node.keywords:
                        if keyword.arg == "mode":
                            mode = str(_literal_path(keyword.value) or "")
                target = self.spec.outputs if method in WRITE_METHODS or (method == "open" and any(flag in mode for flag in "wax+")) else self.spec.inputs
                target.add(_normalise(owner_path))

        if node.args:
            argument = _literal_path(node.args[0])
            if argument:
                if name in READ_FUNCTIONS or name.startswith(("pandas.read_", "pd.read_")):
                    self.spec.inputs.add(_normalise(argument))
                elif method in {"save", "savefig", "imwrite"}:
                    self.spec.outputs.add(_normalise(argument))
                elif name == "open":
                    mode = str(_literal_path(node.args[1]) or "") if len(node.args) > 1 else "r"
                    target = self.spec.outputs if any(flag in mode for flag in "wax+") else self.spec.inputs
                    target.add(_normalise(argument))
        self.generic_visit(node)


def discover_python_artifacts(source: str) -> PythonArtifactSpec:
    """Return literal file inputs/outputs visible in a Python syntax tree."""
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return PythonArtifactSpec()
    visitor = _Visitor()
    visitor.visit(tree)
    visitor.spec.inputs -= visitor.spec.outputs
    return visitor.spec
