"""Tests for ``tools/lint_tools.py`` -- the MCP-facing lint/fit tools.

Boundary validation, envelopes, filters, and response caps; the engine
itself is covered by test_lint_deviant / test_lint_conformant /
test_lint_rules_unit.
"""

import asyncio
import json
import re
from pathlib import Path

import pytest

from tests.conftest import (
    HOUSE_CORPUS_DIR,
    house_corpus_missing,
)
from tools.lint_tools import MAX_LINT_RESPONSE_BYTES, register_lint_tools

DEVIANT = str(HOUSE_CORPUS_DIR / "deviant_01.pptx")
HOUSE_01 = str(HOUSE_CORPUS_DIR / "house_01.pptx")
HOUSE_02 = str(HOUSE_CORPUS_DIR / "house_02.pptx")

needs_corpus = pytest.mark.skipif(
    house_corpus_missing(), reason="house corpus not present")


class _RecorderApp:
    """Minimal stand-in for FastMCP capturing registered tool fns."""

    def __init__(self):
        self.tools = {}

    def tool(self, *args, **kwargs):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator


@pytest.fixture
def tools():
    app = _RecorderApp()
    register_lint_tools(app)
    return app.tools


@pytest.fixture
def loaded_profile(house_profile):
    """The live meridian profile parked in the style-tools registry."""
    from tools import style_tools

    style_tools._style_profiles["meridian_lint_test"] = house_profile
    yield "meridian_lint_test"
    style_tools._style_profiles.pop("meridian_lint_test", None)


# ------------------------------------------------------------ registration

def test_all_three_tools_register(tools):
    assert set(tools) == {"lint_against_profile", "predict_text_fit",
                          "diff_decks"}


def test_tools_register_on_real_fastmcp_app():
    from mcp.server.fastmcp import FastMCP

    app = FastMCP(name="lint-tools-test")
    register_lint_tools(app)
    tool_names = {tool.name for tool in asyncio.run(app.list_tools())}
    assert {"lint_against_profile", "predict_text_fit",
            "diff_decks"} <= tool_names


# ------------------------------------------------------------ lint tool

def test_lint_unknown_profile_is_actionable(tools):
    result = tools["lint_against_profile"]("whatever.pptx", "nope")
    assert "not found" in result["error"]


def test_lint_non_house_profile_is_rejected(tools):
    from tools import style_tools

    style_tools._style_profiles["flat_legacy"] = {"name": "flat"}
    try:
        result = tools["lint_against_profile"]("whatever.pptx",
                                               "flat_legacy")
        assert "house-profile/1" in result["error"]
    finally:
        style_tools._style_profiles.pop("flat_legacy", None)


@needs_corpus
def test_lint_missing_deck(tools, loaded_profile):
    result = tools["lint_against_profile"]("no/such.pptx",
                                           loaded_profile)
    assert "not found" in result["error"]


@needs_corpus
def test_lint_envelope_and_summary(tools, loaded_profile):
    result = tools["lint_against_profile"](DEVIANT, loaded_profile)
    assert "error" not in result
    assert result["summary"]["total"] == len(result["findings"]) == 9
    assert result["summary"]["by_severity"]["error"] == 9
    assert set(result["summary"]["by_rule"]) == {
        "font-scale", "off-grid", "bullet-style", "hardcoded-color",
        "straggler-textbox", "spacing", "border-style", "color-palette",
        "font-family",
    }
    assert result["truncated"] is False
    payload = json.dumps(result, ensure_ascii=False)
    assert len(payload) <= MAX_LINT_RESPONSE_BYTES


@needs_corpus
def test_lint_severity_floor_and_slide_range(tools, loaded_profile):
    errors_only = tools["lint_against_profile"](
        DEVIANT, loaded_profile, severity_floor="error")
    assert all(f["severity"] == "error"
               for f in errors_only["findings"])

    slide_two = tools["lint_against_profile"](
        DEVIANT, loaded_profile, slide_range="2")
    assert {f["slide"] for f in slide_two["findings"]} == {2}
    assert slide_two["summary"]["total"] == 4  # v1, v3, v6, v9

    bad = tools["lint_against_profile"](DEVIANT, loaded_profile,
                                        severity_floor="fatal")
    assert "severity_floor" in bad["error"]

    bad_range = tools["lint_against_profile"](DEVIANT, loaded_profile,
                                              slide_range="7-9")
    assert "error" in bad_range


@needs_corpus
def test_lint_response_capped_with_marker(tools, loaded_profile,
                                          monkeypatch):
    import tools.lint_tools as module

    monkeypatch.setattr(module, "MAX_LINT_RESPONSE_BYTES", 2_000)
    result = tools["lint_against_profile"](DEVIANT, loaded_profile)
    assert result["truncated"] is True
    assert "hint" in result
    assert len(result["findings"]) < 9
    assert result["summary"]["total"] == 9  # summary keeps full counts
    assert len(json.dumps(result, ensure_ascii=False)) <= 2_000


# ------------------------------------------------------------ text fit

def test_predict_missing_deck(tools):
    result = tools["predict_text_fit"]("no/such.pptx")
    assert "not found" in result["error"]


@needs_corpus
def test_predict_envelope(tools):
    result = tools["predict_text_fit"](HOUSE_01, slide_index=0)
    assert "error" not in result
    assert result["frames"], "title slide has text frames"
    assert all(f["slide"] == 1 for f in result["frames"])
    assert sum(result["summary"].values()) == len(result["frames"])


@needs_corpus
def test_predict_bad_slide_index(tools):
    result = tools["predict_text_fit"](HOUSE_01, slide_index=99)
    assert "error" in result


# ------------------------------------------------------------ diff_decks

@needs_corpus
def test_diff_decks_runs_deck_b_as_profile(tools):
    result = tools["diff_decks"](HOUSE_01, HOUSE_02)
    assert "error" not in result
    assert result["reference_deck"] == HOUSE_02
    assert isinstance(result["findings"], list)
    assert result["summary"]["total"] == len(result["findings"])
    # Same house template on both sides: the deterministic style rules
    # (fonts, sizes, colors, spacing, bullets, borders) must agree.
    # Geometry conventions learned from ONE deck are weaker (grid lines
    # need min-slide support), so only geometry rules may fire.
    style_rules = {"font-scale", "font-family", "bullet-style",
                   "spacing", "color-palette", "hardcoded-color",
                   "border-style"}
    style_errors = [f for f in result["findings"]
                    if f["rule_id"] in style_rules]
    assert style_errors == []


@needs_corpus
def test_diff_decks_missing_deck(tools):
    result = tools["diff_decks"](HOUSE_01, "no/such.pptx")
    assert "not found" in result["error"]


# ------------------------------------------------------------ docs

_ROOT = Path(__file__).resolve().parents[1]


def test_readme_documents_lint_tools_and_count():
    readme = (_ROOT / "README.md").read_text(encoding="utf-8")
    for name in ("lint_against_profile", "predict_text_fit",
                 "diff_decks"):
        assert name in readme, f"README must document {name}"
    assert "53" in readme, "README headline tool count must be 53"


def test_readme_module_count_matches_tools_package():
    """Regression: README still said '13 organized modules' in two
    places after the 14th tool module (tools/lint_tools.py) landed.
    Every 'organized ... modules' mention must track the real
    tools/ package contents."""
    actual = len(list((_ROOT / "tools").glob("*_tools.py")))
    readme = (_ROOT / "README.md").read_text(encoding="utf-8")
    mentions = re.findall(r"(\d+) organized (?:tool )?modules", readme)
    assert mentions, "README must mention the organized module count"
    for count in mentions:
        assert int(count) == actual, (
            f"README claims {count} organized modules but tools/ "
            f"contains {actual} *_tools.py modules")


def test_notice_attributes_pptx_lint():
    notice = (_ROOT / "NOTICE").read_text(encoding="utf-8")
    assert "pptx-lint" in notice
