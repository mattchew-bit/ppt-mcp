"""Shared pytest helpers for the ppt-mcp test suite.

Fixture decks live in ``tests/fixtures/`` and are authored in desktop
PowerPoint -- never generated with python-pptx (which would test the
library against itself). Hand-recorded expected values live in
``tests/fixtures/expected_values/<name>.json``.

Conventions:
    * Slide indices in expected-values JSON are 0-based, matching
      python-pptx (``presentation.slides[0]`` is the first slide).
    * Tests that depend on a fixture deck must skip -- not fail -- when
      the deck has not been generated yet (see ``skip_if_fixture_missing``).
"""

import json
from pathlib import Path

import pytest

TESTS_DIR = Path(__file__).resolve().parent
FIXTURES_DIR = TESTS_DIR / "fixtures"
EXPECTED_VALUES_DIR = FIXTURES_DIR / "expected_values"

# The four Step 0 fixture decks (see the style-fidelity-upgrade plan):
#   theme_only        -- values inherited straight from the theme
#   layout_override   -- master/layout-level overrides
#   explicit_override -- explicit run-level overrides (corporate-template style)
#   multi_master      -- two masters with deliberately different themes
FIXTURE_DECKS = (
    "theme_only.pptx",
    "layout_override.pptx",
    "explicit_override.pptx",
    "multi_master.pptx",
)


def fixture_path(name: str) -> Path:
    """Absolute path of a fixture file inside ``tests/fixtures/``."""
    return FIXTURES_DIR / name


def load_expected(name: str) -> dict:
    """Load ``tests/fixtures/expected_values/<name>.json``.

    ``name`` is the fixture stem without extension, e.g. ``"theme_only"``.
    Raises ``FileNotFoundError`` / ``json.JSONDecodeError`` on bad fixtures.
    """
    path = EXPECTED_VALUES_DIR / f"{name}.json"
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def fixture_missing(name: str) -> bool:
    """True when a fixture file has not been generated yet."""
    return not fixture_path(name).is_file()


def skip_if_fixture_missing(name: str):
    """``skipif`` marker for tests that depend on a not-yet-authored fixture."""
    return pytest.mark.skipif(
        fixture_missing(name),
        reason=f"fixture {name} not present in tests/fixtures/ (not generated yet)",
    )


# ---------------------------------------------------------------------------
# Step 3 house corpus (tests/fixtures/house_corpus/)
# ---------------------------------------------------------------------------

HOUSE_CORPUS_DIR = FIXTURES_DIR / "house_corpus"

#: The five Meridian house decks the profile engine learns from.
HOUSE_CORPUS_DECKS = tuple(f"house_{i:02d}.pptx" for i in range(1, 6))


def house_corpus_paths() -> list:
    """Absolute paths of the five house-corpus decks (as strings)."""
    return [str(HOUSE_CORPUS_DIR / name) for name in HOUSE_CORPUS_DECKS]


def house_corpus_missing() -> bool:
    """True when any house-corpus deck has not been generated yet."""
    return any(not (HOUSE_CORPUS_DIR / name).is_file()
               for name in HOUSE_CORPUS_DECKS)


def skip_if_house_corpus_missing():
    """``skipif`` marker for tests needing the full house corpus."""
    return pytest.mark.skipif(
        house_corpus_missing(),
        reason="house corpus not present in tests/fixtures/house_corpus/",
    )


def load_corpus_truth() -> dict:
    """The seeded-truth metadata for the house corpus."""
    path = HOUSE_CORPUS_DIR / "corpus_truth.json"
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


@pytest.fixture(scope="session")
def house_profile():
    """Session-cached house profile built from the five corpus decks."""
    if house_corpus_missing():
        pytest.skip("house corpus not present in tests/fixtures/house_corpus/")
    from utils.profile_extract import create_house_profile

    return create_house_profile(house_corpus_paths(), "meridian_test")


@pytest.fixture(scope="session")
def house_facts():
    """Session-cached resolved slide facts for the five corpus decks."""
    if house_corpus_missing():
        pytest.skip("house corpus not present in tests/fixtures/house_corpus/")
    from utils.profile_extract import collect_corpus_facts

    return collect_corpus_facts(house_corpus_paths())


@pytest.fixture
def open_deck():
    """Factory fixture: open a fixture deck via python-pptx.

    Usage::

        def test_something(open_deck):
            presentation = open_deck("theme_only.pptx")

    Skips the calling test when the fixture deck is missing.
    """
    from pptx import Presentation

    def _open(name: str):
        path = fixture_path(name)
        if not path.is_file():
            pytest.skip(f"fixture {name} not present in tests/fixtures/")
        return Presentation(str(path))

    return _open
