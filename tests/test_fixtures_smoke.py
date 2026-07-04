"""Smoke tests for the Step 0 PowerPoint-authored fixture decks.

Each of the four fixture decks must:
    * open via ``pptx.Presentation`` without raising,
    * contain at least 3 slides,
    * ship a parsable expected-values JSON whose slide references all
      exist in the deck (0-based indices, matching python-pptx).

Tests skip -- not fail -- while a fixture deck has not been generated yet
(fixtures are authored in desktop PowerPoint, separately from this code).
"""

from pathlib import Path

import pytest

from tests.conftest import (
    EXPECTED_VALUES_DIR,
    FIXTURE_DECKS,
    load_expected,
    skip_if_fixture_missing,
)

_SLIDE_INDEX_KEYS = frozenset({"slide_index", "slide_idx", "slide"})


def _deck_params():
    """One pytest param per fixture deck, skipping decks not yet authored."""
    return [
        pytest.param(deck, marks=skip_if_fixture_missing(deck), id=deck)
        for deck in FIXTURE_DECKS
    ]


def referenced_slide_indices(data) -> frozenset:
    """Collect every 0-based slide index referenced by an expected-values JSON.

    Recognized shapes (tolerant of schema details, strict about indices):
        * any dict key in ``{"slide_index", "slide_idx", "slide"}`` whose
          value is an integer,
        * a ``"slides"`` dict keyed by numeric strings (or ints),
        * a ``"slides"`` list, whose positions imply indices 0..len-1.
    """
    found = set()
    _walk(data, found)
    return frozenset(found)


def _walk(node, found) -> None:
    if isinstance(node, dict):
        for key, value in node.items():
            is_int = isinstance(value, int) and not isinstance(value, bool)
            if key in _SLIDE_INDEX_KEYS and is_int:
                found.add(value)
            elif key == "slides" and isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    text = str(sub_key)
                    if text.lstrip("-").isdigit():
                        found.add(int(text))
                    _walk(sub_value, found)
            elif key == "slides" and isinstance(value, list):
                found.update(range(len(value)))
                _walk(value, found)
            else:
                _walk(value, found)
    elif isinstance(node, list):
        for item in node:
            _walk(item, found)


@pytest.mark.parametrize("deck_name", _deck_params())
def test_deck_opens_without_exception(deck_name, open_deck):
    presentation = open_deck(deck_name)
    assert presentation is not None


@pytest.mark.parametrize("deck_name", _deck_params())
def test_deck_has_at_least_three_slides(deck_name, open_deck):
    presentation = open_deck(deck_name)
    assert len(presentation.slides) >= 3, (
        f"{deck_name} has {len(presentation.slides)} slides; "
        "Step 0 fixtures require at least 3"
    )


@pytest.mark.parametrize("deck_name", _deck_params())
def test_expected_values_json_exists_and_references_valid_slides(
    deck_name, open_deck
):
    stem = Path(deck_name).stem
    json_path = EXPECTED_VALUES_DIR / f"{stem}.json"
    assert json_path.is_file(), (
        f"missing expected-values JSON for {deck_name}: {json_path}"
    )

    expected = load_expected(stem)  # raises JSONDecodeError if unparsable
    assert expected, f"{json_path} parsed to an empty document"

    presentation = open_deck(deck_name)
    slide_count = len(presentation.slides)
    out_of_range = sorted(
        index
        for index in referenced_slide_indices(expected)
        if not 0 <= index < slide_count
    )
    assert not out_of_range, (
        f"{json_path} references slide indices {out_of_range} but "
        f"{deck_name} only has slides 0..{slide_count - 1} "
        "(indices are 0-based)"
    )
