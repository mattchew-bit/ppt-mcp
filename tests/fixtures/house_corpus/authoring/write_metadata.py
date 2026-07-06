"""Write corpus_truth.json + deviations.json from the authoring constants.

The truth file records EVERY seeded convention of the Meridian corpus --
theme palette, type scale, master text styles, paragraph metrics,
bullets, shape defaults, the 3-column grid, per-archetype geometry,
footer/image zones, per-slide archetype labels and the theme-color tint
seeds -- straight from ``house_style`` / ``author_house_decks`` /
``author_deviant``, so it cannot drift from the decks.

Expected tint hexes are computed with the repo's own
``utils.resolve_colors.apply_color_transforms``; ``transform_check.py``
then arbitrates them against PowerPoint-COM-reported effective values.
"""

from __future__ import annotations

import datetime as _dt
import json

from lxml import etree

import _bootstrap
import house_style as hs
from author_deviant import DECK_NAME as DEVIANT_NAME
from author_deviant import DEVIATIONS, SLIDE_LABELS
from author_house_decks import DECKS
from utils.resolve_colors import apply_color_transforms

TRUTH_PATH = _bootstrap.corpus_path("corpus_truth.json")
DEVIATIONS_PATH = _bootstrap.corpus_path("deviations.json")

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _rect(rect_pt: tuple) -> dict:
    """(x, y, w, h) points -> dict carrying both pt and in values."""
    x, y, w, h = rect_pt
    return {
        "x_pt": x, "y_pt": y, "w_pt": w, "h_pt": h,
        "x_in": round(x / 72.0, 4), "y_in": round(y / 72.0, 4),
        "w_in": round(w / 72.0, 4), "h_in": round(h / 72.0, 4),
    }


def _edges(values_pt: tuple) -> dict:
    return {"pt": list(values_pt),
            "in": [round(v / 72.0, 4) for v in values_pt]}


def expected_tint_hex(token: str, brightness: float) -> str:
    """Predict the effective hex of a theme-color brightness variant."""
    transforms = hs.brightness_transforms(brightness)
    elements = [
        etree.fromstring(f'<a:{name} xmlns:a="{_A_NS}" val="{val}"/>')
        for name, val in transforms.items()
    ]
    return apply_color_transforms(hs.SCHEME_HEX[token], elements)


def tinted_elements() -> list[dict]:
    """Walk the deck specs and list every theme-color tint seed."""
    records: list[dict] = []
    for deck_name, deck in DECKS.items():
        for index, spec in enumerate(deck["slides"], start=1):
            tint = spec.get("tint")
            if not tint:
                continue
            targets = {
                "rule": ("AccentRule", "fill"),
                "kicker": ("SectionKicker", "font"),
                "panels": (("ColumnPanelLeft", "ColumnPanelRight"), "fill"),
            }
            for key, seed in tint.items():
                shapes, target = targets[key]
                if isinstance(shapes, str):
                    shapes = (shapes,)
                for shape_name in shapes:
                    records.append({
                        "deck": f"{deck_name}.pptx", "slide": index,
                        "shape": shape_name, "target": target,
                        "scheme_token": seed["token"],
                        "base_hex": hs.SCHEME_HEX[seed["token"]],
                        "brightness": seed["brightness"],
                        "xml_transforms":
                            hs.brightness_transforms(seed["brightness"]),
                        "expected_hex":
                            expected_tint_hex(seed["token"],
                                              seed["brightness"]),
                    })
    return records


def _typography() -> dict:
    body_l1 = hs.BODY_LEVELS[1]
    return {
        "type_scale_pt": list(hs.TYPE_SCALE_PT),
        "title": {"font": hs.HOUSE_MAJOR_FONT, "size_pt": hs.TITLE_SIZE_PT,
                  "bold": hs.TITLE_BOLD, "alignment": hs.TITLE_ALIGNMENT,
                  "color": "#" + hs.SCHEME_HEX[hs.TITLE_COLOR_TOKEN],
                  "color_token": hs.TITLE_COLOR_TOKEN},
        "subtitle": {"font": hs.HOUSE_MINOR_FONT,
                     "size_pt": hs.SUBTITLE_SIZE_PT,
                     "color": "#" + hs.SCHEME_HEX["accent3"],
                     "color_token": "accent3"},
        "body": {"font": hs.HOUSE_MINOR_FONT,
                 "size_pt": body_l1["size_pt"],
                 "color": "#" + hs.SCHEME_HEX[body_l1["color_token"]],
                 "color_token": body_l1["color_token"]},
        "footer": {"font": hs.HOUSE_MINOR_FONT,
                   "size_pt": hs.CAPTION_SIZE_PT,
                   "color": "#" + hs.SCHEME_HEX["accent3"],
                   "color_token": "accent3"},
        "kicker": {"font": hs.HOUSE_MINOR_FONT,
                   "size_pt": hs.CAPTION_SIZE_PT, "bold": True,
                   "color": "#" + hs.SCHEME_HEX["accent1"],
                   "color_token": "accent1"},
        "caption": {"font": hs.HOUSE_MINOR_FONT,
                    "size_pt": hs.CAPTION_SIZE_PT, "italic": True,
                    "color": "#" + hs.SCHEME_HEX["accent3"],
                    "color_token": "accent3"},
        "panel_text": {"font": hs.HOUSE_MINOR_FONT,
                       "size_pt": hs.BODY_SIZE_PT,
                       "header_bold": True,
                       "color": "#" + hs.SCHEME_HEX["dk1"],
                       "color_token": "dk1"},
    }


def _paragraph_levels() -> dict:
    levels = {}
    for index, seed in hs.BODY_LEVELS.items():
        levels[str(index)] = {
            "size_pt": seed["size_pt"],
            "space_before_pt": seed["space_before_pt"],
            "space_after_pt": seed["space_after_pt"],
            "line_spacing_multiple": seed["line_spacing"],
            "bullet": {"char": seed["bullet_glyph"],
                       "char_code": seed["bullet_char"],
                       "font": seed["bullet_font"],
                       "size_pct": round(seed["bullet_rel_size"] * 100)},
        }
    return levels


def _archetypes() -> dict:
    geometry = {
        "title": {"title_band": _rect(hs.TITLE_MAIN),
                  "body_region": _rect(hs.TITLE_SUBTITLE),
                  "extras": {"rule": _rect(hs.TITLE_RULE)}},
        "agenda": {"title_band": _rect(hs.TITLE_BAND),
                   "body_region": _rect(hs.BODY_AGENDA),
                   "extras": {"image": _rect(hs.SIDEBAR_IMAGE),
                              "caption": _rect(hs.SIDEBAR_CAPTION)}},
        "section_divider": {"title_band": _rect(hs.DIVIDER_TITLE),
                            "body_region": _rect(hs.DIVIDER_RULE),
                            "extras": {"kicker": _rect(hs.DIVIDER_KICKER)}},
        "content": {"title_band": _rect(hs.TITLE_BAND),
                    "body_region": _rect(hs.BODY_FULL),
                    "extras": {"takeaway": _rect(hs.TAKEAWAY_PANEL)}},
        "two_column": {"title_band": _rect(hs.TITLE_BAND),
                       "body_region": _rect((60.0, 120.0, 540.0, 300.0)),
                       "extras": {"panel_left": _rect(hs.COLUMN_PANEL_LEFT),
                                  "panel_right":
                                      _rect(hs.COLUMN_PANEL_RIGHT),
                                  "image": _rect(hs.SIDEBAR_IMAGE),
                                  "caption": _rect(hs.SIDEBAR_CAPTION)}},
        "closing": {"title_band": _rect(hs.CLOSING_TITLE),
                    "body_region": _rect(hs.CONTACT_PANEL),
                    "extras": {"rule": _rect(hs.CLOSING_RULE)}},
    }
    labels = _slide_labels()
    for name, record in geometry.items():
        slides = [f"{deck}:{idx}" for deck, deck_labels in labels.items()
                  for idx, label in enumerate(deck_labels, start=1)
                  if label == name]
        record["count"] = len(slides)
        record["slides"] = slides
    return geometry


def _slide_labels() -> dict[str, list[str]]:
    return {name: [spec["archetype"] for spec in deck["slides"]]
            for name, deck in DECKS.items()}


def _image_slides() -> list[str]:
    return [f"{deck}:{idx}" for deck, deck_labels in _slide_labels().items()
            for idx, label in enumerate(deck_labels, start=1)
            if label in ("agenda", "two_column")]


def build_truth() -> dict:
    labels = _slide_labels()
    return {
        "schema": "house-corpus-truth/1",
        "generated_by": "tests/fixtures/house_corpus/authoring/"
                        "write_metadata.py",
        "generated_on": _dt.date.today().isoformat(),
        "template": {"name": "Meridian", "master": hs.HOUSE_MASTER_NAME},
        "slide_size": {"width_pt": hs.SLIDE_W_PT, "height_pt": hs.SLIDE_H_PT,
                       "width_in": round(hs.SLIDE_W_PT / 72.0, 4),
                       "height_in": round(hs.SLIDE_H_PT / 72.0, 4)},
        "theme": {"scheme": {token: "#" + value
                             for token, value in hs.SCHEME_HEX.items()},
                  "fonts": {"major_latin": hs.HOUSE_MAJOR_FONT,
                            "minor_latin": hs.HOUSE_MINOR_FONT}},
        "typography": _typography(),
        "paragraph": {"scope": "body placeholder levels (master bodyStyle)",
                      "levels": _paragraph_levels()},
        "shape_defaults": {
            "border": {"weight_pt": hs.BORDER_WEIGHT_PT,
                       "color": "#" + hs.BORDER_COLOR_HEX,
                       "color_token": "dk2",
                       "dash": hs.BORDER_DASH_NAME},
            "corner_radius_adj": hs.CORNER_RADIUS_ADJ,
            "fill": "#" + hs.PANEL_FILL_HEX,
            "fill_token": "lt2",
            "applies_to": "content panels (rounded rectangles)"},
        "grid": {"columns": 3,
                 "left_edges": _edges(hs.GRID_LEFT_EDGES_PT),
                 "right_edges": _edges(hs.GRID_RIGHT_EDGES_PT),
                 "center_edges": _edges(hs.GRID_CENTER_EDGES_PT),
                 "column_width_pt": hs.GRID_COLUMN_W_PT,
                 "gutter_pt": hs.GRID_GUTTER_PT,
                 "tolerance_pt": hs.GRID_TOLERANCE_PT,
                 "tolerance_in": round(hs.GRID_TOLERANCE_PT / 72.0, 4)},
        "archetypes": _archetypes(),
        "footer": {"zone": _rect(hs.FOOTER_ZONE),
                   "note_box": _rect(hs.FOOTER_NOTE),
                   "page_box": _rect(hs.FOOTER_PAGE),
                   "note_text": hs.FOOTER_NOTE_TEXT,
                   "applies_to": list(hs.FOOTER_ARCHETYPES)},
        "images": {"zones": {"sidebar": _rect(hs.SIDEBAR_IMAGE)},
                   "size_pt": {"width": hs.SIDEBAR_IMAGE[2],
                               "height": hs.SIDEBAR_IMAGE[3]},
                   "assets": ["authoring/assets/exhibit_bars.png",
                              "authoring/assets/exhibit_blocks.png",
                              "authoring/assets/exhibit_wave.png"],
                   "slides_with_images": _image_slides()},
        "decks": {f"{name}.pptx": {"title": deck["title_text"],
                                   "slide_count": len(deck["slides"]),
                                   "slides": labels[name]}
                  for name, deck in DECKS.items()},
        "labeled_slide_count": sum(len(v) for v in labels.values()),
        "tinted_elements": tinted_elements(),
    }


def build_deviations() -> dict:
    return {
        "schema": "house-corpus-deviations/1",
        "generated_by": "tests/fixtures/house_corpus/authoring/"
                        "write_metadata.py",
        "deck": f"{DEVIANT_NAME}.pptx",
        "template": {"name": "Meridian", "master": hs.HOUSE_MASTER_NAME},
        "slides": list(SLIDE_LABELS),
        "violation_count": len(DEVIATIONS),
        "violations": DEVIATIONS,
    }


def _write(path: str, document: dict) -> str:
    with open(path, "w", encoding="utf-8", newline="\n") as handle:
        json.dump(document, handle, indent=2, ensure_ascii=False)
        handle.write("\n")
    print(f"Wrote {path}")
    return path


def main() -> tuple[str, str]:
    return (_write(TRUTH_PATH, build_truth()),
            _write(DEVIATIONS_PATH, build_deviations()))


if __name__ == "__main__":
    main()
