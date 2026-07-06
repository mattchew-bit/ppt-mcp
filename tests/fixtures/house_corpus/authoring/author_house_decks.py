"""Author the five Meridian house decks via PowerPoint COM.

27 labeled slides across house_01..house_05 covering all six archetypes
(title x5, agenda x4, section_divider x4, content x6, two_column x5,
closing x3). Three slides carry theme-color tints (lumMod/lumOff
variants): house_01 slide 3 (rule fill + kicker font), house_01 slide 5
(both column panel fills) and house_02 slide 5 (rule fill) -- brightness
values chosen from the COM probe sweep where PowerPoint's rounding is
stable (not on a .5 channel boundary).

``DECKS`` is data, consumed both here (to author) and by
``write_metadata.py`` (to record labels + tints in corpus_truth.json).
"""

from __future__ import annotations

import _bootstrap
from com_helpers import new_presentation, powerpoint_app, save_pptx

from house_style import apply_house_master
from slide_archetypes import build_slide

DECK_NAMES = ("house_01", "house_02", "house_03", "house_04", "house_05")

DECKS: dict[str, dict] = {
    "house_01": {
        "title_text": "Meridian market entry review",
        "slides": [
            {"archetype": "title",
             "title": "Meridian market entry review",
             "subtitle": "Prepared for the executive committee"},
            {"archetype": "agenda",
             "title": "Agenda for the working session",
             "items": [(1, "Where the market stands today"),
                       (1, "Entry routes on the shortlist"),
                       (1, "Investment ask and phasing"),
                       (1, "Decision points for this committee"),
                       (1, "Next steps and owners")],
             "image": "bars",
             "caption": "Exhibit 1 - category growth, synthetic"},
            {"archetype": "section_divider",
             "title": "Where the market stands",
             "kicker": "SECTION 01",
             "tint": {"rule": {"token": "accent1", "brightness": 0.4},
                      "kicker": {"token": "accent2", "brightness": -0.25}}},
            {"archetype": "content",
             "title": "Demand is recovering faster than supply",
             "bullets": [(1, "Category demand grew nine percent this year"),
                         (2, "Premium tier carried most of the growth"),
                         (2, "Entry tier stayed flat for four quarters"),
                         (1, "Supply additions lag by roughly two years"),
                         (2, "Two announced plants slipped to next cycle"),
                         (3, "Permitting drove both slips"),
                         (1, "Pricing held firm across all regions")],
             "takeaway": ["A two year demand-supply gap keeps pricing "
                          "attractive for a new entrant"]},
            {"archetype": "two_column",
             "title": "Two entry routes lead the shortlist",
             "left": {"header": "Build",
                      "lines": ["Full control of the asset",
                                "Slowest route to first revenue",
                                "Highest capital at risk"]},
             "right": {"header": "Partner",
                       "lines": ["Shared economics with incumbent",
                                 "Fastest route to first revenue",
                                 "Limited influence on roadmap"]},
             "image": "blocks",
             "caption": "Exhibit 2 - route comparison, synthetic",
             "tint": {"panels": {"token": "accent1", "brightness": 0.8}}},
            {"archetype": "closing",
             "title": "Thank you",
             "contact": {"header": "Contact",
                         "lines": ["Meridian Advisory",
                                   "strategy team, synthetic corpus",
                                   "meridian.example"]}},
        ],
    },
    "house_02": {
        "title_text": "Operating model diagnostic",
        "slides": [
            {"archetype": "title",
             "title": "Operating model diagnostic",
             "subtitle": "Readout for the transformation office"},
            {"archetype": "agenda",
             "title": "What we will cover today",
             "items": [(1, "Diagnostic scope and method"),
                       (1, "Findings across the four functions"),
                       (1, "Cost and service level trade-offs"),
                       (1, "Recommended sequence of moves")],
             "image": "wave",
             "caption": "Exhibit 1 - service trend, synthetic"},
            {"archetype": "content",
             "title": "Handoffs between functions drive most delay",
             "bullets": [(1, "Order journey crosses five functional lines"),
                         (2, "Each handoff adds queue time"),
                         (2, "No single owner tracks the full journey"),
                         (1, "Rework concentrates in two steps"),
                         (2, "Both steps rely on manual re-entry"),
                         (1, "Automation covers under a third of volume")],
             "takeaway": ["Collapsing two handoffs removes roughly half "
                          "of the end to end delay"]},
            {"archetype": "two_column",
             "title": "Centralize or federate the shared services",
             "left": {"header": "Centralize",
                      "lines": ["One global process owner",
                                "Deepest scale economics",
                                "Longer change program"]},
             "right": {"header": "Federate",
                       "lines": ["Regional autonomy preserved",
                                 "Faster local adoption",
                                 "Duplicated tooling remains"]},
             "image": "blocks",
             "caption": "Exhibit 2 - model options, synthetic"},
            {"archetype": "section_divider",
             "title": "Recommended sequence",
             "kicker": "SECTION 03",
             "tint": {"rule": {"token": "accent1", "brightness": 0.6}}},
            {"archetype": "closing",
             "title": "Thank you",
             "contact": {"header": "Contact",
                         "lines": ["Meridian Advisory",
                                   "operations practice, synthetic",
                                   "meridian.example"]}},
        ],
    },
    "house_03": {
        "title_text": "Pricing program update",
        "slides": [
            {"archetype": "title",
             "title": "Pricing program update",
             "subtitle": "Monthly steering committee readout"},
            {"archetype": "agenda",
             "title": "Program status at a glance",
             "items": [(1, "Wave one results against target"),
                       (1, "Wave two pilot readiness"),
                       (1, "Risks and mitigations"),
                       (1, "Asks of the steering committee")],
             "image": "bars",
             "caption": "Exhibit 1 - realized uplift, synthetic"},
            {"archetype": "content",
             "title": "Wave one delivered above the uplift target",
             "bullets": [(1, "Realized uplift reached the top of the range"),
                         (2, "List price moves landed with low churn"),
                         (2, "Discount discipline held in both regions"),
                         (1, "Win rates stayed inside the guardrail"),
                         (3, "Two accounts required manual exceptions")],
             "takeaway": ["Wave one economics validate the playbook for "
                          "the next two waves"]},
            {"archetype": "content",
             "title": "Wave two pilots start in two markets",
             "bullets": [(1, "Pilot markets chosen for contract mix"),
                         (2, "Renewal-heavy book in the first market"),
                         (2, "New-logo-heavy book in the second"),
                         (1, "Playbook adapted for channel partners"),
                         (2, "Partner margin floor stays unchanged")],
             "takeaway": ["Pilot design isolates the two contract "
                          "archetypes before full rollout"]},
            {"archetype": "two_column",
             "title": "Escalate or hold the exception policy",
             "left": {"header": "Hold policy",
                      "lines": ["Preserves price integrity",
                                "Risks two flagship renewals",
                                "Simple to communicate"]},
             "right": {"header": "Escalate",
                       "lines": ["Retains the flagship logos",
                                 "Signals flexibility to the field",
                                 "Requires deal desk review"]},
             "image": "wave",
             "caption": "Exhibit 2 - exception volume, synthetic"},
        ],
    },
    "house_04": {
        "title_text": "Supply chain resilience assessment",
        "slides": [
            {"archetype": "title",
             "title": "Supply chain resilience assessment",
             "subtitle": "Findings and countermeasure options"},
            {"archetype": "section_divider",
             "title": "Exposure by tier",
             "kicker": "SECTION 02"},
            {"archetype": "content",
             "title": "Tier two concentration is the binding risk",
             "bullets": [(1, "Three inputs share a single tier two source"),
                         (2, "All three sit in one coastal cluster"),
                         (1, "Qualified alternates exist for one input"),
                         (2, "Qualification lead time is nine months"),
                         (1, "Inventory covers six weeks of disruption")],
             "takeaway": ["Dual-sourcing the two unqualified inputs closes "
                          "most of the exposure"]},
            {"archetype": "two_column",
             "title": "Buffer stock versus dual sourcing",
             "left": {"header": "Buffer stock",
                      "lines": ["Live within one quarter",
                                "Ties up working capital",
                                "Decays as demand shifts"]},
             "right": {"header": "Dual source",
                       "lines": ["Nine month qualification",
                                 "Durable structural fix",
                                 "Small unit cost premium"]},
             "image": "blocks",
             "caption": "Exhibit 1 - option economics, synthetic"},
            {"archetype": "closing",
             "title": "Thank you",
             "contact": {"header": "Contact",
                         "lines": ["Meridian Advisory",
                                   "supply chain practice, synthetic",
                                   "meridian.example"]}},
        ],
    },
    "house_05": {
        "title_text": "Digital roadmap briefing",
        "slides": [
            {"archetype": "title",
             "title": "Digital roadmap briefing",
             "subtitle": "Quarterly portfolio review"},
            {"archetype": "agenda",
             "title": "Portfolio review agenda",
             "items": [(1, "Delivery status by initiative"),
                       (1, "Budget consumption against plan"),
                       (1, "Re-prioritization proposals"),
                       (1, "Decisions requested today")],
             "image": "wave",
             "caption": "Exhibit 1 - delivery velocity, synthetic"},
            {"archetype": "content",
             "title": "Two initiatives need a scope decision",
             "bullets": [(1, "Customer portal is ahead of schedule"),
                         (2, "Second release can pull forward"),
                         (1, "Data platform burn exceeds plan"),
                         (2, "Scope grew after the pilot feedback"),
                         (3, "Storage costs doubled the estimate"),
                         (1, "Field app awaits the platform decision")],
             "takeaway": ["Trimming platform scope funds the portal "
                          "pull-forward inside the same envelope"]},
            {"archetype": "two_column",
             "title": "Trim scope or extend the timeline",
             "left": {"header": "Trim scope",
                      "lines": ["Holds the budget envelope",
                                "Defers two data domains",
                                "Keeps the release date"]},
             "right": {"header": "Extend timeline",
                       "lines": ["Delivers the full scope",
                                 "Adds two quarters of burn",
                                 "Delays dependent projects"]},
             "image": "bars",
             "caption": "Exhibit 2 - scope options, synthetic"},
            {"archetype": "section_divider",
             "title": "Decisions requested",
             "kicker": "SECTION 04"},
        ],
    },
}


def build_deck(app, deck_name: str) -> str:
    """Author one house deck and save it into the corpus directory."""
    spec = DECKS[deck_name]
    output = _bootstrap.corpus_path(f"{deck_name}.pptx")
    with new_presentation(app) as pres:
        master = pres.Designs(1).SlideMaster
        apply_house_master(master)
        for page_no, slide_spec in enumerate(spec["slides"], start=1):
            build_slide(pres, master, slide_spec, page_no)
        return save_pptx(pres, output)


def main() -> list[str]:
    written = []
    with powerpoint_app() as app:
        for deck_name in DECK_NAMES:
            path = build_deck(app, deck_name)
            written.append(path)
            print(f"Wrote {path}")
    return written


if __name__ == "__main__":
    main()
