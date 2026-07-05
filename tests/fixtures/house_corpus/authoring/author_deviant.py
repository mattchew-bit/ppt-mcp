"""Author deviant_01.pptx: the SAME Meridian template with 9 seeded
style violations, for apply/lint testing.

Every violation is recorded in ``DEVIATIONS`` (single source -- it is
serialized verbatim into ``deviations.json`` by write_metadata.py).
Apart from the seeded violations the deck is house-conformant, so a
linter must flag exactly these and nothing else.

Violations:
    v1 font_size_offscale   13pt run (house scale 11/14/20/30)
    v2 off_grid_shape       panel left edge 682pt (col3 edge 660, +22pt)
    v3 wrong_bullet_char    l1 bullet u2022 (house l1 is em dash u2014)
    v4 hardcoded_srgb       run colored via literal RGB -> a:srgbClr
                            20262B where the house uses schemeClr dk1
    v5 footer_straggler     stray text box inside the footer zone
    v6 wrong_space_after    paragraph space_after 12pt (house 8pt)
    v7 wrong_border_weight  takeaway panel border 2.5pt (house 1.25pt)
    v8 off_palette_color    panel fill 8E44AD (not in the house scheme)
    v9 wrong_font           run in Times New Roman (house body Calibri)
"""

from __future__ import annotations

import _bootstrap
from com_helpers import (
    add_textbox,
    new_presentation,
    paragraph,
    powerpoint_app,
    save_pptx,
    set_font,
    set_paragraph,
    solid_fill,
)

import house_style as hs
from house_style import apply_house_master
from slide_archetypes import build_slide

DECK_NAME = "deviant_01"

SLIDE_LABELS = ("title", "content", "content", "two_column", "closing")

#: Off-grid seed: nearest house edge is col3 left (660pt); offset +22pt.
OFF_GRID_RECT = (682.0, 150.0, 240.0, 120.0)
STRAGGLER_RECT = (400.0, 505.0, 200.0, 18.0)

DEVIATIONS: list[dict] = [
    {"id": "v1", "kind": "font_size_offscale", "slide": 2,
     "shape": "BodyContent", "paragraph": 2, "property": "font.size_pt",
     "expected": 14.0, "actual": 13.0,
     "note": "house type scale is {11, 14, 20, 30}pt"},
    {"id": "v2", "kind": "off_grid_shape", "slide": 3,
     "shape": "OffGridPanel", "property": "geometry.left_pt",
     "expected": 660.0, "actual": 682.0,
     "note": "22pt right of the column 3 left edge (tolerance 4pt)"},
    {"id": "v3", "kind": "wrong_bullet_char", "slide": 2,
     "shape": "BodyContent", "paragraph": 4, "property": "bullet.char",
     "expected": "—", "actual": "•",
     "note": "house level 1 bullet is the em dash"},
    {"id": "v4", "kind": "hardcoded_srgb", "slide": 4,
     "shape": "ColumnPanelLeft", "paragraph": 1,
     "property": "font.color_source",
     "expected": "schemeClr dk1", "actual": "srgbClr 20262B",
     "note": "same visual hex but hardcoded instead of theme-linked"},
    {"id": "v5", "kind": "footer_straggler", "slide": 5,
     "shape": "StragglerNote", "property": "geometry",
     "expected": "no stray shapes in footer zone y 500-528pt",
     "actual": "text box at (400, 505, 200, 18)pt",
     "note": "also off-grid; footer zone allows FooterNote/FooterPage only"},
    {"id": "v6", "kind": "wrong_space_after", "slide": 2,
     "shape": "BodyContent", "paragraph": 1,
     "property": "paragraph.space_after_pt",
     "expected": 8.0, "actual": 12.0,
     "note": "house body level 1 space_after is 8pt"},
    {"id": "v7", "kind": "wrong_border_weight", "slide": 3,
     "shape": "TakeawayPanel", "property": "line.weight_pt",
     "expected": 1.25, "actual": 2.5,
     "note": "house panel border is 1.25pt"},
    {"id": "v8", "kind": "off_palette_color", "slide": 4,
     "shape": "ColumnPanelRight", "property": "fill.color_hex",
     "expected": "DCE3E8", "actual": "8E44AD",
     "note": "8E44AD is not one of the 12 scheme colors"},
    {"id": "v9", "kind": "wrong_font", "slide": 2,
     "shape": "BodyContent", "paragraph": 3, "property": "font.name",
     "expected": "Calibri", "actual": "Times New Roman",
     "note": "house body font is the theme minor font (Calibri)"},
]

_SLIDE_SPECS: list[dict] = [
    {"archetype": "title",
     "title": "Regional expansion options",
     "subtitle": "Working draft for internal review"},
    {"archetype": "content",
     "title": "Volume outlook supports a second line",
     "bullets": [(1, "Committed volume fills the current line"),
                 (1, "Spot demand is turned away every month"),
                 (1, "Two anchor customers asked for capacity"),
                 (1, "Competitor lead times keep lengthening")],
     "takeaway": ["Demand evidence supports the expansion case"]},
    {"archetype": "content",
     "title": "Site shortlist narrows to two regions",
     "bullets": [(1, "Northern site has the stronger labor pool"),
                 (2, "Training partner is already on site"),
                 (1, "Southern site has the cheaper energy"),
                 (2, "Grid interconnect is pre-approved")],
     "takeaway": ["Both sites clear the hurdle rate; labor risk "
                  "separates them"]},
    {"archetype": "two_column",
     "title": "Lease the shell or build to suit",
     "left": {"header": "Lease shell",
              "lines": ["Occupancy in two quarters",
                        "Landlord controls expansion",
                        "Lower upfront cash"]},
     "right": {"header": "Build to suit",
               "lines": ["Occupancy in six quarters",
                         "Full layout control",
                         "Higher upfront cash"]},
     "image": "blocks",
     "caption": "Exhibit 1 - site options, synthetic"},
    {"archetype": "closing",
     "title": "Thank you",
     "contact": {"header": "Contact",
                 "lines": ["Meridian Advisory",
                           "expansion team, synthetic",
                           "meridian.example"]}},
]


def _body_paragraph(slide, index: int):
    """Paragraph range *index* (1-based) of the BodyContent placeholder."""
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name == "BodyContent":
            return paragraph(shape.TextFrame2.TextRange, index)
    raise LookupError("slide has no BodyContent placeholder")


def _find_shape(slide, name: str):
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == name:
            return slide.Shapes(i)
    raise LookupError(f"shape {name!r} not found")


def _seed_slide2(slide) -> None:
    set_paragraph(_body_paragraph(slide, 1), space_after_pt=12.0)   # v6
    set_font(_body_paragraph(slide, 2), size=13.0)                  # v1
    set_font(_body_paragraph(slide, 3), name="Times New Roman")     # v9
    set_paragraph(_body_paragraph(slide, 4), bullet_char=0x2022)    # v3


def _seed_slide3(slide) -> None:
    hs.add_panel(slide, OFF_GRID_RECT, "OffGridPanel",               # v2
                 "Side note", ["Late addition, pasted in"])
    takeaway = _find_shape(slide, "TakeawayPanel")
    takeaway.Line.Weight = 2.5                                       # v7


def _seed_slide4(slide) -> None:
    left = _find_shape(slide, "ColumnPanelLeft")
    header = paragraph(left.TextFrame2.TextRange, 1)
    set_font(header, color_hex=hs.SCHEME_HEX["dk1"])                 # v4
    right = _find_shape(slide, "ColumnPanelRight")
    solid_fill(right, "8E44AD")                                      # v8


def _seed_slide5(slide) -> None:
    box = add_textbox(slide, *STRAGGLER_RECT,                        # v5
                      "Draft - not for distribution",
                      name="StragglerNote")
    set_font(box.TextFrame2.TextRange, size=hs.CAPTION_SIZE_PT)


def build_deck(app) -> str:
    output = _bootstrap.corpus_path(f"{DECK_NAME}.pptx")
    with new_presentation(app) as pres:
        master = pres.Designs(1).SlideMaster
        apply_house_master(master)
        slides = []
        for page_no, spec in enumerate(_SLIDE_SPECS, start=1):
            slides.append(build_slide(pres, master, spec, page_no))
        _seed_slide2(slides[1])
        _seed_slide3(slides[2])
        _seed_slide4(slides[3])
        _seed_slide5(slides[4])
        return save_pptx(pres, output)


def main() -> str:
    with powerpoint_app() as app:
        path = build_deck(app)
        print(f"Wrote {path} ({len(DEVIATIONS)} seeded deviations)")
        return path


if __name__ == "__main__":
    main()
