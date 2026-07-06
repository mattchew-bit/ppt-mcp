"""Slide builders for the six house archetypes.

Each builder consumes one slide *spec* dict (see
``author_house_decks.DECKS``) and authors a slide whose every shape edge
snaps to the Meridian grid. Tint specs apply PowerPoint theme-color
brightness variants (lumMod/lumOff in the XML) -- the raw material for
the Step 2 transform-math fixture check.

Spec keys by archetype:
    title           title, subtitle
    agenda          title, items [(level, text)...], image (asset key)
    section_divider title, kicker, tint? {rule: {token, brightness},
                    kicker: {token, brightness}}
    content         title, bullets [(level, text)...], takeaway [lines]
    two_column      title, left {header, lines}, right {header, lines},
                    image, caption?, tint? {panels: {token, brightness}}
    closing         title, contact {header, lines}
"""

from __future__ import annotations

import _bootstrap
from com_helpers import add_slide, find_layout, set_bullet_text

import house_style as hs

#: Asset key -> committed PNG filename (see make_images.ASSETS).
IMAGE_ASSETS = {
    "bars": "exhibit_bars.png",
    "blocks": "exhibit_blocks.png",
    "wave": "exhibit_wave.png",
}


def _asset(key: str) -> str:
    if key not in IMAGE_ASSETS:
        raise KeyError(f"unknown image asset key: {key!r}")
    return _bootstrap.asset_path(IMAGE_ASSETS[key])


def build_title(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title Only"))
    hs.add_rule(slide, hs.TITLE_RULE, name="TitleRule")
    hs.place_title(slide, spec["title"], rect=hs.TITLE_MAIN)
    hs.add_subtitle(slide, spec["subtitle"])
    return slide


def build_agenda(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title and Content"))
    hs.place_title(slide, spec["title"])
    body = hs.place_body(slide, hs.BODY_AGENDA)
    set_bullet_text(body, spec["items"])
    hs.add_image(slide, _asset(spec["image"]))
    hs.add_caption(slide, spec.get("caption", "Exhibit - synthetic data"))
    hs.add_footer(slide, page_no)
    return slide


def build_section_divider(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title Only"))
    kicker = hs.add_kicker(slide, spec["kicker"])
    hs.place_title(slide, spec["title"], rect=hs.DIVIDER_TITLE)
    rule = hs.add_rule(slide, hs.DIVIDER_RULE)
    tint = spec.get("tint", {})
    if "rule" in tint:
        hs.tint_fill(rule, tint["rule"]["token"], tint["rule"]["brightness"])
    if "kicker" in tint:
        hs.tint_font(kicker.TextFrame2.TextRange,
                     tint["kicker"]["token"], tint["kicker"]["brightness"])
    return slide


def build_content(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title and Content"))
    hs.place_title(slide, spec["title"])
    body = hs.place_body(slide, hs.BODY_FULL)
    set_bullet_text(body, spec["bullets"])
    hs.add_panel(slide, hs.TAKEAWAY_PANEL, "TakeawayPanel",
                 "Key takeaway", spec["takeaway"])
    hs.add_footer(slide, page_no)
    return slide


def build_two_column(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title Only"))
    hs.place_title(slide, spec["title"])
    left = hs.add_panel(slide, hs.COLUMN_PANEL_LEFT, "ColumnPanelLeft",
                        spec["left"]["header"], spec["left"]["lines"])
    right = hs.add_panel(slide, hs.COLUMN_PANEL_RIGHT, "ColumnPanelRight",
                         spec["right"]["header"], spec["right"]["lines"])
    tint = spec.get("tint", {})
    if "panels" in tint:
        for panel in (left, right):
            hs.tint_fill(panel, tint["panels"]["token"],
                         tint["panels"]["brightness"])
    hs.add_image(slide, _asset(spec["image"]))
    hs.add_caption(slide, spec.get("caption", "Exhibit - synthetic data"))
    hs.add_footer(slide, page_no)
    return slide


def build_closing(pres, master, spec: dict, page_no: int):
    slide = add_slide(pres, find_layout(master, "Title Only"))
    hs.add_rule(slide, hs.CLOSING_RULE, name="ClosingRule")
    hs.place_title(slide, spec["title"], rect=hs.CLOSING_TITLE)
    hs.add_panel(slide, hs.CONTACT_PANEL, "ContactPanel",
                 spec["contact"]["header"], spec["contact"]["lines"])
    hs.add_footer(slide, page_no)
    return slide


BUILDERS = {
    "title": build_title,
    "agenda": build_agenda,
    "section_divider": build_section_divider,
    "content": build_content,
    "two_column": build_two_column,
    "closing": build_closing,
}


def build_slide(pres, master, spec: dict, page_no: int):
    """Dispatch one spec to its archetype builder."""
    archetype = spec["archetype"]
    if archetype not in BUILDERS:
        raise KeyError(f"unknown archetype: {archetype!r}")
    return BUILDERS[archetype](pres, master, spec, page_no)
