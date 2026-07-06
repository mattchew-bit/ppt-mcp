"""Step 5 gate (a): lint the deviant corpus deck against the LIVE-built
Meridian house profile.

``deviant_01.pptx`` carries exactly nine seeded style violations,
recorded in ``tests/fixtures/house_corpus/deviations.json`` (the seed
corpus the plan names). This suite asserts:

* every recorded violation is flagged by the CORRECT rule (rule_id is
  asserted per violation) with the correct expected/actual values;
* the deck draws NO OTHER error/warn findings -- apart from the seeded
  violations it is house-conformant, so a linter must flag exactly
  these and nothing else (author_deviant.py contract);
* findings come back ordered by severity then slide.

These tests were written FIRST (TDD) -- they are the executable spec
for ``utils.lint_engine`` / ``utils.lint_rules``.
"""

import json

import pytest

from tests.conftest import HOUSE_CORPUS_DIR, house_corpus_missing

DEVIANT = HOUSE_CORPUS_DIR / "deviant_01.pptx"

pytestmark = pytest.mark.skipif(
    house_corpus_missing() or not DEVIANT.is_file(),
    reason="house corpus / deviant deck not present",
)


def _deviations():
    with (HOUSE_CORPUS_DIR / "deviations.json").open(
            "r", encoding="utf-8") as handle:
        return json.load(handle)


@pytest.fixture(scope="module")
def findings(house_profile):
    from utils.lint_engine import lint_against_profile

    return lint_against_profile(str(DEVIANT), house_profile)


def _norm_hex(value):
    return str(value).lstrip("#").upper()


def _hex_pool(value):
    """expected/actual color fields may be a hex string or a list."""
    if isinstance(value, (list, tuple)):
        return {_norm_hex(v) for v in value}
    return {_norm_hex(value)}


def _match(findings, **keys):
    """Findings matching every given key exactly."""
    out = []
    for finding in findings:
        if all(finding.get(k) == v for k, v in keys.items()):
            out.append(finding)
    return out


def _one(findings, vid, **keys):
    hits = _match(findings, **keys)
    assert len(hits) == 1, (
        f"{vid}: expected exactly one finding matching {keys}, got "
        f"{len(hits)}: {hits!r}"
    )
    return hits[0]


#: violation id -> the rule that must flag it (the plan's rule catalog).
RULE_BY_VIOLATION = {
    "v1": "font-scale",
    "v2": "off-grid",
    "v3": "bullet-style",
    "v4": "hardcoded-color",
    "v5": "straggler-textbox",
    "v6": "spacing",
    "v7": "border-style",
    "v8": "color-palette",
    "v9": "font-family",
}


def test_every_recorded_violation_has_a_rule_mapping():
    recorded = {v["id"] for v in _deviations()["violations"]}
    assert recorded == set(RULE_BY_VIOLATION)


def test_v1_font_size_offscale(findings):
    f = _one(findings, "v1", rule_id="font-scale", slide=2,
             shape="BodyContent", paragraph=2)
    assert f["severity"] == "error"
    assert f["property"] == "font.size_pt"
    assert f["actual"] == pytest.approx(13.0)
    assert 14.0 in f["expected"], "house scale must appear in expected"
    # Distribution-style message: the house type scale, not bare fail.
    for member in ("11", "14", "20", "30"):
        assert member in f["message"]


def test_v2_off_grid_shape(findings):
    f = _one(findings, "v2", rule_id="off-grid", slide=3,
             shape="OffGridPanel")
    assert f["severity"] == "error"
    assert f["property"] == "geometry.left_pt"
    assert f["actual"] == pytest.approx(682.0, abs=0.3)
    # Profile grid edges are rounded to 2 decimals (9.17in = 660.24pt).
    assert f["expected"] == pytest.approx(660.0, abs=0.5)
    # Distance-to-nearest-gridline appears in the message.
    assert "21.8" in f["message"] or "22" in f["message"]


def test_v3_wrong_bullet_char(findings):
    f = _one(findings, "v3", rule_id="bullet-style", slide=2,
             shape="BodyContent", paragraph=4)
    assert f["severity"] == "error"
    assert f["property"] == "bullet.char"
    assert f["expected"] == "—"   # em dash
    assert f["actual"] == "•"     # bullet


def test_v4_hardcoded_srgb(findings):
    f = _one(findings, "v4", rule_id="hardcoded-color", slide=4,
             shape="ColumnPanelLeft", paragraph=1)
    assert f["severity"] == "error"
    assert f["property"] == "font.color_source"
    assert f["expected"] == "schemeClr dk1"
    assert f["actual"] == "srgbClr 20262B"


def test_v5_footer_straggler(findings):
    f = _one(findings, "v5", rule_id="straggler-textbox", slide=5,
             shape="StragglerNote")
    assert f["severity"] == "error"
    assert f["property"] == "geometry"
    assert "400" in str(f["actual"]) and "505" in str(f["actual"])


def test_v6_wrong_space_after(findings):
    f = _one(findings, "v6", rule_id="spacing", slide=2,
             shape="BodyContent", paragraph=1)
    assert f["severity"] == "error"
    assert f["property"] == "paragraph.space_after_pt"
    assert f["expected"] == pytest.approx(8.0)
    assert f["actual"] == pytest.approx(12.0)
    # Distribution-style message: the learned space_after quanta.
    assert "8" in f["message"]


def test_v7_wrong_border_weight(findings):
    f = _one(findings, "v7", rule_id="border-style", slide=3,
             shape="TakeawayPanel")
    assert f["severity"] == "error"
    assert f["property"] == "line.weight_pt"
    assert f["expected"] == pytest.approx(1.25)
    assert f["actual"] == pytest.approx(2.5)


def test_v8_off_palette_color(findings):
    f = _one(findings, "v8", rule_id="color-palette", slide=4,
             shape="ColumnPanelRight")
    assert f["severity"] == "error"
    assert f["property"] == "fill.color_hex"
    assert _hex_pool(f["actual"]) == {"8E44AD"}
    # The house panel fill (deviations.json expected) is in the allowed
    # palette the finding reports.
    assert "DCE3E8" in _hex_pool(f["expected"])


def test_v9_wrong_font(findings):
    f = _one(findings, "v9", rule_id="font-family", slide=2,
             shape="BodyContent", paragraph=3)
    assert f["severity"] == "error"
    assert f["property"] == "font.name"
    assert f["actual"] == "Times New Roman"
    expected = (f["expected"] if isinstance(f["expected"], (list, tuple))
                else [f["expected"]])
    assert "Calibri" in expected


def test_no_extra_error_or_warn_findings(findings):
    """The deviant deck is conformant apart from the nine seeds."""
    seeded = {
        (RULE_BY_VIOLATION[v["id"]], v["slide"], v["shape"])
        for v in _deviations()["violations"]
    }
    flagged = [
        (f["rule_id"], f["slide"], f["shape"])
        for f in findings if f["severity"] in ("error", "warn")
    ]
    extras = [key for key in flagged if key not in seeded]
    assert not extras, f"unexpected error/warn findings: {extras!r}"
    assert len(flagged) == len(seeded) == 9


def test_findings_ordered_by_severity_then_slide(findings):
    rank = {"error": 0, "warn": 1, "info": 2}
    keys = [(rank[f["severity"]], f["slide"]) for f in findings]
    assert keys == sorted(keys)


def test_findings_are_json_serializable(findings):
    payload = json.dumps(findings, ensure_ascii=False)
    assert isinstance(payload, str) and payload.startswith("[")
