"""Import smoke tests: every ppt-mcp module must import without raising.

Guards the 41 existing MCP tools (zero prior coverage) against regressions
from later style-fidelity steps: if a refactor breaks a module-level import
anywhere under ``tools/`` or ``utils/``, or in the server entry module,
these tests go red immediately.
"""

import importlib
import pkgutil

import pytest


def _package_module_names(package_name: str) -> list:
    """Fully-qualified names of every direct submodule of a package."""
    package = importlib.import_module(package_name)
    return sorted(
        f"{package_name}.{info.name}"
        for info in pkgutil.iter_modules(package.__path__)
    )


def _all_module_names() -> list:
    return (
        ["ppt_mcp_server", "tools", "utils"]
        + _package_module_names("tools")
        + _package_module_names("utils")
    )


@pytest.mark.parametrize("module_name", _all_module_names())
def test_module_imports_cleanly(module_name):
    module = importlib.import_module(module_name)
    assert module is not None
