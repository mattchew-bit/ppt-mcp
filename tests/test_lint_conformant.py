"""Step 5 gate (b): ZERO false positives on the conformant house corpus.

Each of the five Meridian house decks the profile was learned from must
lint completely clean at error/warn level against that same profile --
the zero-FP half of the plan's ISC criterion ("lint flags every seeded
deviation; zero false positives on conformant fixture").

Also covers gate (e): a foreign deck (``theme_only.pptx``, a different
template entirely) lints against the Meridian profile without crashing.
"""

import json

import pytest

from tests.conftest import (
    HOUSE_CORPUS_DECKS,
    HOUSE_CORPUS_DIR,
    fixture_path,
    house_corpus_missing,
)

pytestmark = pytest.mark.skipif(
    house_corpus_missing(),
    reason="house corpus not present in tests/fixtures/house_corpus/",
)


@pytest.mark.parametrize("deck", HOUSE_CORPUS_DECKS)
def test_conformant_deck_has_zero_error_or_warn_findings(deck,
                                                         house_profile):
    from utils.lint_engine import lint_against_profile

    findings = lint_against_profile(str(HOUSE_CORPUS_DIR / deck),
                                    house_profile)
    noisy = [f for f in findings if f["severity"] in ("error", "warn")]
    assert noisy == [], (
        f"{deck}: expected zero error/warn findings on a conformant "
        f"house deck, got {json.dumps(noisy, ensure_ascii=False, indent=1)}"
    )


def test_foreign_deck_lints_without_crashing(house_profile):
    """A deck from a different template must produce findings, not
    exceptions -- the engine's job on foreign content is reporting."""
    from utils.lint_engine import lint_against_profile

    foreign = fixture_path("theme_only.pptx")
    if not foreign.is_file():
        pytest.skip("theme_only.pptx fixture not present")
    findings = lint_against_profile(str(foreign), house_profile)
    assert isinstance(findings, list)
    json.dumps(findings, ensure_ascii=False)  # must stay serializable
    for finding in findings:
        for key in ("rule_id", "severity", "slide", "message"):
            assert key in finding


def test_engine_rejects_non_house_profiles():
    from utils.lint_engine import lint_against_profile

    deck = str(HOUSE_CORPUS_DIR / HOUSE_CORPUS_DECKS[0])
    with pytest.raises(ValueError, match="house-profile/1"):
        lint_against_profile(deck, {"name": "not-a-profile"})


def test_engine_rejects_missing_deck(house_profile):
    from utils.lint_engine import lint_against_profile

    with pytest.raises(FileNotFoundError):
        lint_against_profile("no/such/deck.pptx", house_profile)
