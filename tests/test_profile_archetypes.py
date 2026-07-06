"""Tests for ``utils/profile_archetypes.py`` -- geometric classification.

The Step 3 acceptance criterion: reproduce >= 90% of the corpus's 27
labeled archetypes (corpus_truth.json ``decks.*.slides``). Learned
title-band / body-region boxes are checked against the seeded geometry.
"""

import pytest

from tests.conftest import load_corpus_truth, skip_if_house_corpus_missing
from utils.profile_archetypes import (
    ARCHETYPE_NAMES,
    classify_slide,
    find_title_shape,
    learn_archetypes,
)

#: Body size for the SYNTHETIC unit-test slides below, which author
#: their own 14pt runs (self-consistent truth). Corpus tests must NOT
#: use this constant: they derive the body size from the engine-learned
#: profile so a body-size-learning regression cannot hide behind a
#: constant that happens to equal the seeded truth.
SYNTHETIC_BODY_SIZE_PT = 14.0


def _learned_body_size_pt(house_profile) -> float:
    """The engine-learned body size out of the built profile."""
    return float(house_profile["typography"]["body"]["size"]["value"])

#: Learned boxes may differ from seeded rects by text-autofit height
#: noise (COM shrinks empty textbox heights); positions stay exact.
BOX_TOLERANCE_IN = 0.15


def _shape(left, top, width, height, *, text=None, size=14.0,
           is_placeholder=False, ph_type=None, is_picture=False):
    record = {
        "name": "s",
        "is_placeholder": is_placeholder,
        "ph_type": ph_type,
        "is_picture": is_picture,
        "geometry": {"left_pt": left, "top_pt": top,
                     "width_pt": width, "height_pt": height,
                     "preset": "rect"},
    }
    if text is not None:
        record["paragraphs"] = [{
            "indent_level": 1,
            "runs": [{"text": text,
                      "font": {"size_pt": size, "name": "Calibri",
                               "bold": False, "color_hex": "20262B"}}],
        }]
    return record


def _slide(shapes, deck="synthetic", number=1):
    return {"deck": deck, "slide_number": number,
            "width_pt": 960.0, "height_pt": 540.0, "shapes": shapes}


# ------------------------------------------------------------- unit rules


def test_full_bleed_picture_wins():
    slide = _slide([
        _shape(0, 0, 960, 540, is_picture=True),
        _shape(60, 30, 840, 60, text="Title", size=30,
               is_placeholder=True, ph_type="TITLE (1)"),
    ])
    assert classify_slide(slide, SYNTHETIC_BODY_SIZE_PT) == "full_bleed"


def test_untitled_slide_falls_back_to_content():
    assert classify_slide(_slide([]), SYNTHETIC_BODY_SIZE_PT) == "content"


def test_title_fallback_uses_largest_type_without_placeholder():
    slide = _slide([
        _shape(60, 300, 400, 50, text="body text", size=14),
        _shape(60, 40, 800, 60, text="Big headline", size=30),
    ])
    title = find_title_shape(slide)
    assert title["geometry"]["top_pt"] == 40


def test_classify_rejects_bad_body_size():
    with pytest.raises(ValueError, match="body_size_pt"):
        classify_slide(_slide([]), 0)


def test_learn_archetypes_requires_slides():
    with pytest.raises(ValueError, match="at least one slide"):
        learn_archetypes([], SYNTHETIC_BODY_SIZE_PT)


def test_classification_stays_in_closed_set():
    slide = _slide([
        _shape(60, 30, 840, 60, text="T", size=30,
               is_placeholder=True, ph_type="TITLE (1)"),
    ])
    assert classify_slide(slide, SYNTHETIC_BODY_SIZE_PT) in ARCHETYPE_NAMES


# ------------------------------------------------------ corpus accuracy


@skip_if_house_corpus_missing()
def test_archetype_accuracy_at_least_90_percent(house_facts, house_profile):
    """>= 90% of the 27 labeled corpus slides classify correctly.

    The body size feeding the classifier is the ENGINE-LEARNED one, not
    a constant equal to the seeded truth -- a body-size-learning
    regression must surface here, not hide.
    """
    body_size_pt = _learned_body_size_pt(house_profile)
    truth = load_corpus_truth()
    labels = {
        (deck_name, index + 1): archetype
        for deck_name, deck in truth["decks"].items()
        for index, archetype in enumerate(deck["slides"])
    }
    assert len(labels) == truth["labeled_slide_count"]

    correct, wrong = 0, []
    for slide in house_facts["slides"]:
        expected = labels[(slide["deck"], slide["slide_number"])]
        got = classify_slide(slide, body_size_pt)
        if got == expected:
            correct += 1
        else:
            wrong.append((slide["deck"], slide["slide_number"],
                          expected, got))
    accuracy = correct / len(labels)
    assert accuracy >= 0.9, f"accuracy {accuracy:.2%}; misses: {wrong}"


@skip_if_house_corpus_missing()
def test_every_corpus_archetype_learned_with_correct_count(house_profile):
    truth = load_corpus_truth()["archetypes"]
    learned = house_profile["archetypes"]
    for name, spec in truth.items():
        assert name in learned, f"archetype {name} not learned"
        assert learned[name]["count"] == spec["count"]
    for name in learned:
        assert name in ARCHETYPE_NAMES


@skip_if_house_corpus_missing()
@pytest.mark.parametrize("box", ["title_band", "body_region"])
def test_learned_boxes_match_seeded_geometry(house_profile, box):
    truth = load_corpus_truth()["archetypes"]
    for name, spec in truth.items():
        learned = house_profile["archetypes"][name][box]
        for learned_key, truth_key in (("x", "x_in"), ("y", "y_in"),
                                       ("w", "w_in"), ("h", "h_in")):
            leaf = learned[learned_key]
            assert leaf["unit"] == "in"
            delta = abs(leaf["value"] - spec[box][truth_key])
            assert delta <= BOX_TOLERANCE_IN, (
                f"{name}.{box}.{learned_key}: learned {leaf['value']} "
                f"vs seeded {spec[box][truth_key]}"
            )


@skip_if_house_corpus_missing()
def test_title_bands_are_tight(house_profile):
    """Title placeholders are set exactly; allow only rounding slack."""
    truth = load_corpus_truth()["archetypes"]
    for name, spec in truth.items():
        learned = house_profile["archetypes"][name]["title_band"]
        for learned_key, truth_key in (("x", "x_in"), ("y", "y_in"),
                                       ("w", "w_in"), ("h", "h_in")):
            delta = abs(learned[learned_key]["value"]
                        - spec["title_band"][truth_key])
            assert delta <= 0.05, f"{name}.title_band.{learned_key}"
