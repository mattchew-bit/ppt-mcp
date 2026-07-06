"""End-to-end tests for the Step 3 house-profile engine.

Asserts the profile built from the five Meridian corpus decks against
the SEEDED truth (corpus_truth.json) -- exact values, never just
non-None -- plus the pinned schema contract: DTCG leaf shapes, the 8KB
size budget, byte-identical determinism, DTCG export, persist/load
round-trip through the existing storage mechanics, and the MCP tool.
"""

import json

import pytest

from tests.conftest import (
    house_corpus_paths,
    load_corpus_truth,
    skip_if_house_corpus_missing,
)
from utils.profile_schema import (
    MAX_PROFILE_BYTES,
    SCHEMA_VERSION,
    enforce_size_budget,
    serialize_profile,
    to_dtcg,
)

# ------------------------------------------------------------- helpers


def _hex(hex_no_hash: str) -> str:
    return f"#{hex_no_hash.lstrip('#').upper()}"


# --------------------------------------------------------- input errors


def test_create_house_profile_rejects_empty_paths():
    from utils.profile_extract import create_house_profile

    with pytest.raises(ValueError, match="non-empty"):
        create_house_profile([], "x")


def test_create_house_profile_rejects_missing_deck():
    from utils.profile_extract import create_house_profile

    with pytest.raises(FileNotFoundError, match="no_such_deck"):
        create_house_profile(["no_such_deck.pptx"], "x")


def test_create_house_profile_rejects_bad_name():
    from utils.profile_extract import create_house_profile

    with pytest.raises(ValueError, match="name"):
        create_house_profile(["a.pptx"], "")


# ------------------------------------------------------------ top level


@skip_if_house_corpus_missing()
def test_profile_top_level_contract(house_profile):
    assert house_profile["schema_version"] == SCHEMA_VERSION
    assert house_profile["name"] == "meridian_test"
    assert house_profile["source_decks"] == [
        f"house_{i:02d}.pptx" for i in range(1, 6)
    ]
    expected_keys = {
        "schema_version", "name", "source_decks", "slide_size",
        "typography", "paragraph", "palette", "shape_defaults",
        "grid", "archetypes", "images", "distributions",
    }
    assert set(house_profile) == expected_keys


@skip_if_house_corpus_missing()
def test_slide_size_matches_truth(house_profile):
    truth = load_corpus_truth()["slide_size"]
    size = house_profile["slide_size"]
    assert size["width"] == {"value": round(truth["width_in"], 2),
                             "unit": "in"}
    assert size["height"] == {"value": round(truth["height_in"], 2),
                              "unit": "in"}


# ------------------------------------------------------------ typography


@skip_if_house_corpus_missing()
def test_typography_title_matches_seeded_master_style(house_profile):
    truth = load_corpus_truth()["typography"]["title"]
    title = house_profile["typography"]["title"]
    assert title["font"] == {"value": truth["font"]}
    assert title["size"] == {"value": truth["size_pt"], "unit": "pt"}
    assert title["bold"] == {"value": truth["bold"]}
    assert title["color"] == {"value": _hex(truth["color"])}


@skip_if_house_corpus_missing()
def test_typography_body_matches_seeded_master_style(house_profile):
    truth = load_corpus_truth()["typography"]["body"]
    body = house_profile["typography"]["body"]
    assert body["font"] == {"value": truth["font"]}
    assert body["size"] == {"value": truth["size_pt"], "unit": "pt"}
    assert body["color"] == {"value": _hex(truth["color"])}


@skip_if_house_corpus_missing()
def test_typography_footer_matches_seeded_convention(house_profile):
    truth = load_corpus_truth()["typography"]["footer"]
    footer = house_profile["typography"]["footer"]
    assert footer["font"] == {"value": truth["font"]}
    assert footer["size"] == {"value": truth["size_pt"], "unit": "pt"}
    assert footer["color"] == {"value": _hex(truth["color"])}


# ------------------------------------------------------------- paragraph


@skip_if_house_corpus_missing()
def test_paragraph_spacing_matches_seeded_body_style(house_profile):
    truth = load_corpus_truth()["paragraph"]["levels"]["1"]
    paragraph = house_profile["paragraph"]
    assert paragraph["space_before"] == {
        "value": truth["space_before_pt"], "unit": "pt"}
    assert paragraph["space_after"] == {
        "value": truth["space_after_pt"], "unit": "pt"}
    assert paragraph["line_spacing"] == {
        "value": truth["line_spacing_multiple"], "unit": None}


@skip_if_house_corpus_missing()
@pytest.mark.parametrize("level", ["1", "2", "3"])
def test_bullet_rules_match_seeded_levels(house_profile, level):
    truth = load_corpus_truth()["paragraph"]["levels"][level]["bullet"]
    rule = house_profile["paragraph"]["bullets"][f"l{level}"]
    assert rule["char"] == {"value": truth["char"]}
    assert rule["size_pct"] == {"value": float(truth["size_pct"]),
                                "unit": None}


@skip_if_house_corpus_missing()
def test_bullet_rules_carry_only_pinned_contract_keys(house_profile):
    """Regression: the pinned v1 contract is {char, color?, size_pct?}.

    The builder once widened it with a ``font`` key -- contract drift
    that no consumer reads. Every emitted bullet rule must stay inside
    the pinned key set.
    """
    bullets = house_profile["paragraph"]["bullets"]
    assert bullets, "corpus should learn at least one bullet level"
    for level, rule in bullets.items():
        assert set(rule) <= {"char", "color", "size_pct"}, level


# --------------------------------------------------------------- palette


@skip_if_house_corpus_missing()
def test_palette_scheme_matches_seeded_theme(house_profile):
    truth = load_corpus_truth()["theme"]["scheme"]
    scheme = house_profile["palette"]["scheme"]
    assert len(scheme) == 12
    for name, hex_value in truth.items():
        assert scheme[name] == {"value": _hex(hex_value)}


@skip_if_house_corpus_missing()
def test_palette_usage_shares_are_sane(house_profile):
    usage = house_profile["palette"]["usage"]
    assert 0 < len(usage) <= 8
    assert all(0.0 <= entry["share"] <= 1.0 for entry in usage)
    assert sum(entry["share"] for entry in usage) <= 1.01
    shares = [entry["share"] for entry in usage]
    assert shares == sorted(shares, reverse=True)
    roles = {entry["color"]: entry["role"] for entry in usage}
    assert roles[_hex("DCE3E8")] == "fill"      # panel fill convention
    assert roles[_hex("20262B")] == "text"      # dk1 body text
    assert roles[_hex("14324F")] == "text"      # dk2 titles


# --------------------------------------------------------- shape defaults


@skip_if_house_corpus_missing()
def test_shape_defaults_match_seeded_panel_conventions(house_profile):
    truth = load_corpus_truth()["shape_defaults"]
    defaults = house_profile["shape_defaults"]
    assert defaults["border"]["weight"] == {
        "value": truth["border"]["weight_pt"], "unit": "pt"}
    assert defaults["border"]["color"] == {
        "value": _hex(truth["border"]["color"])}
    assert defaults["border"]["dash"] == {"value": truth["border"]["dash"]}
    assert defaults["corner_radius"] == {
        "value": truth["corner_radius_adj"], "unit": None}
    assert defaults["fill"] == {"value": _hex(truth["fill"])}


# ---------------------------------------------------------------- images


@skip_if_house_corpus_missing()
def test_image_stats_match_seeded_placements(house_profile):
    truth = load_corpus_truth()
    images = house_profile["images"]
    image_slides = len(truth["images"]["slides_with_images"])
    total_slides = truth["labeled_slide_count"]
    assert images["count_per_slide"]["max"] == 1
    assert images["count_per_slide"]["mean"] == round(
        image_slides / total_slides, 2)

    dominant = images["size_distribution"][0]
    assert dominant["width"]["value"] == pytest.approx(
        truth["images"]["size_pt"]["width"] / 72.0, abs=0.02)
    assert dominant["height"]["value"] == pytest.approx(
        truth["images"]["size_pt"]["height"] / 72.0, abs=0.02)
    assert dominant["share"] == 1.0

    zone = images["zones"][0]
    seeded = truth["images"]["zones"]["sidebar"]
    for key, seeded_key in (("x", "x_in"), ("y", "y_in"),
                            ("w", "w_in"), ("h", "h_in")):
        assert zone[key]["value"] == pytest.approx(
            seeded[seeded_key], abs=0.05)
    assert zone["share"] == 1.0


# ---------------------------------------------------------- distributions


@skip_if_house_corpus_missing()
def test_font_size_distribution_recovers_type_scale(house_profile):
    truth = load_corpus_truth()["typography"]["type_scale_pt"]
    distribution = house_profile["distributions"]["font_sizes_pt"]
    assert distribution["values"] == truth  # exactly {11, 14, 20, 30}
    assert sum(distribution["shares"]) == pytest.approx(1.0, abs=0.05)


@skip_if_house_corpus_missing()
def test_space_after_distribution_contains_seeded_quanta(house_profile):
    truth_levels = load_corpus_truth()["paragraph"]["levels"]
    values = house_profile["distributions"]["space_after_pt"]["values"]
    for level in truth_levels.values():
        assert level["space_after_pt"] in values


@skip_if_house_corpus_missing()
def test_palette_shares_distribution(house_profile):
    distribution = house_profile["distributions"]["palette_shares"]
    assert len(distribution["values"]) == len(distribution["shares"])
    assert distribution["values"][0] == _hex("20262B")  # dk1 dominates
    assert distribution["shares"] == sorted(distribution["shares"],
                                            reverse=True)


# ------------------------------------------------- schema / DTCG leaves


def _walk_leaves(node, path=""):
    """Yield (path, leaf) for every ``{"value": ...}`` dict."""
    if isinstance(node, dict):
        if "value" in node:
            yield path, node
            return
        for key, child in node.items():
            yield from _walk_leaves(child, f"{path}.{key}")
    elif isinstance(node, list):
        for index, child in enumerate(node):
            yield from _walk_leaves(child, f"{path}[{index}]")


@skip_if_house_corpus_missing()
def test_dtcg_leaf_shapes(house_profile):
    sections = ("slide_size", "typography", "paragraph", "shape_defaults",
                "archetypes")
    seen = 0
    for section in sections:
        for path, leaf in _walk_leaves(house_profile[section], section):
            seen += 1
            assert set(leaf) <= {"value", "unit"}, path
            if "unit" in leaf:
                assert leaf["unit"] in ("pt", "in", None), path
                assert isinstance(leaf["value"], (int, float)), path
            value = leaf["value"]
            if isinstance(value, str) and value.startswith("#"):
                assert len(value) == 7 and value == value.upper(), path
    assert seen > 40  # the profile is leaf-shaped, not accidentally flat


@skip_if_house_corpus_missing()
def test_archetype_rects_are_dtcg_inch_leaves(house_profile):
    for name, spec in house_profile["archetypes"].items():
        for box in ("title_band", "body_region"):
            rect = spec[box]
            assert set(rect) == {"x", "y", "w", "h"}, name
            for leaf in rect.values():
                assert leaf["unit"] == "in"


# ------------------------------------------------------------ size budget


@skip_if_house_corpus_missing()
def test_profile_fits_size_budget(house_profile):
    payload = enforce_size_budget(house_profile)
    assert len(payload) <= MAX_PROFILE_BYTES


@skip_if_house_corpus_missing()
def test_size_budget_violation_names_heavy_sections(house_profile):
    bloated = dict(house_profile)
    bloated["distributions"] = {
        "font_sizes_pt": {"values": list(range(3000)), "shares": []},
        "space_after_pt": {"values": [], "shares": []},
        "palette_shares": {"values": [], "shares": []},
    }
    with pytest.raises(ValueError) as excinfo:
        enforce_size_budget(bloated)
    message = str(excinfo.value)
    assert "over the 8192-byte budget" in message
    assert "distributions=" in message  # heaviest section is named


def test_serialize_rejects_foreign_dicts():
    with pytest.raises(ValueError, match="schema_version"):
        serialize_profile({"name": "not-a-profile"})


# ----------------------------------------------------------- determinism


@skip_if_house_corpus_missing()
def test_profile_build_is_deterministic(house_profile):
    """Same input -> byte-identical serialized profile."""
    from utils.profile_extract import create_house_profile

    rebuilt = create_house_profile(house_corpus_paths(), "meridian_test")
    assert serialize_profile(rebuilt) == serialize_profile(house_profile)


# ----------------------------------------------------------- DTCG export


@skip_if_house_corpus_missing()
def test_to_dtcg_colors_and_typography(house_profile):
    truth = load_corpus_truth()
    tokens = to_dtcg(house_profile)

    scheme = tokens["color"]["scheme"]
    for name, hex_value in truth["theme"]["scheme"].items():
        assert scheme[name] == {"$type": "color",
                                "$value": _hex(hex_value)}

    title = tokens["typography"]["title"]
    assert title["$type"] == "typography"
    assert title["$value"]["fontFamily"] == truth["typography"]["title"]["font"]
    # 30pt -> 40px at CSS density (96/72)
    assert title["$value"]["fontSize"] == {"value": 40.0, "unit": "px"}
    assert title["$value"]["fontWeight"] == 400

    assert tokens["color"]["text"]["title"]["$value"] == _hex(
        truth["typography"]["title"]["color"])


def test_to_dtcg_rejects_foreign_dicts():
    with pytest.raises(ValueError, match="house-profile"):
        to_dtcg({"schema_version": "something-else"})


def test_to_dtcg_tolerates_partial_typography_roles():
    """Regression: the builder emits only modal-FOUND keys, so a role
    can lack font/size/bold (e.g. a footer whose corpus runs resolve no
    font names or sizes). Such a partial spec is a VALID profile and
    must not crash the exporter (it used to raise a bare KeyError); the
    role is simply inexpressible as a DTCG typography composite and is
    skipped, while its text color still exports."""
    profile = {
        "schema_version": SCHEMA_VERSION,
        "name": "partial_roles",
        "typography": {
            "title": {
                "font": {"value": "Georgia"},
                "size": {"value": 30.0, "unit": "pt"},
                "bold": {"value": False},
                "color": {"value": "#14324F"},
            },
            # Partial: no resolvable font/size in the corpus.
            "footer": {
                "bold": {"value": False},
                "color": {"value": "#3E5C76"},
            },
            # Partial without even a color: skipped everywhere.
            "body": {"bold": {"value": True}},
        },
        "palette": {"scheme": {"dk1": {"value": "#20262B"}}},
    }
    tokens = to_dtcg(profile)
    assert set(tokens["typography"]) == {"title"}
    assert tokens["typography"]["title"]["$value"]["fontFamily"] == "Georgia"
    assert tokens["color"]["text"]["footer"]["$value"] == "#3E5C76"
    assert "body" not in tokens["color"]["text"]


# ------------------------------------------------- persist/load round-trip


@skip_if_house_corpus_missing()
def test_save_load_round_trip_is_lossless(house_profile, tmp_path):
    from utils.style_utils import load_profile, save_profile

    target = tmp_path / "house.json"
    save_profile(house_profile, str(target))
    assert target.stat().st_size <= MAX_PROFILE_BYTES

    loaded = load_profile(str(target))
    assert isinstance(loaded, dict)
    assert loaded == house_profile
    assert serialize_profile(loaded) == serialize_profile(house_profile)


@skip_if_house_corpus_missing()
def test_saved_profile_is_valid_compact_json(house_profile, tmp_path):
    from utils.style_utils import save_profile

    target = tmp_path / "house.json"
    save_profile(house_profile, str(target))
    text = target.read_text(encoding="utf-8")
    assert json.loads(text) == house_profile
    assert "\n" not in text  # canonical compact form


# --------------------------------------------------------------- MCP tool


class _RecorderApp:
    """Minimal FastMCP stand-in capturing registered tool functions."""

    def __init__(self):
        self.tools = {}

    def tool(self, *args, **kwargs):
        def decorator(fn):
            self.tools[fn.__name__] = fn
            return fn

        return decorator


def _registered_tools(presentations=None, current_id=None):
    from tools.style_tools import register_style_tools

    app = _RecorderApp()
    register_style_tools(app, presentations or {}, lambda: current_id)
    return app.tools


def test_create_house_profile_tool_registers_on_real_fastmcp():
    import asyncio

    from mcp.server.fastmcp import FastMCP

    from tools.style_tools import register_style_tools

    app = FastMCP(name="profile-tools-test")
    register_style_tools(app, {}, lambda: None)
    tool_names = {tool.name for tool in asyncio.run(app.list_tools())}
    assert "create_house_profile" in tool_names


@skip_if_house_corpus_missing()
def test_create_house_profile_tool_round_trip(tmp_path):
    tools = _registered_tools()

    result = tools["create_house_profile"](
        paths=house_corpus_paths(), profile_name="meridian_tool")
    assert "error" not in result
    assert result["profile_bytes"] <= MAX_PROFILE_BYTES
    assert result["schema_version"] == SCHEMA_VERSION
    assert result["archetypes"]["two_column"] == 5

    target = tmp_path / "meridian_tool.json"
    saved = tools["save_style_profile"](
        profile_name="meridian_tool", output_path=str(target))
    assert "error" not in saved

    loaded = tools["load_style_profile"](file_path=str(target))
    assert "error" not in loaded
    assert loaded["profile_name"] == "meridian_tool"

    fetched = tools["get_style_profile"](profile_name="meridian_tool")
    assert fetched["profile"]["schema_version"] == SCHEMA_VERSION


def test_create_house_profile_tool_reports_errors():
    tools = _registered_tools()
    result = tools["create_house_profile"](paths=[], profile_name="x")
    assert "error" in result
    result = tools["create_house_profile"](
        paths=["definitely_missing.pptx"], profile_name="x")
    assert "error" in result
