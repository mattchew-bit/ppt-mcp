"""The "Meridian" house template: every seeded convention in ONE module.

This is the single source of truth for the Step 3 house corpus. The deck
builders consume these constants when authoring via PowerPoint COM, and
``write_metadata.py`` serializes the very same constants into
``corpus_truth.json`` -- so the truth file can never drift from what the
decks actually contain.

Every value is deliberately NON-DEFAULT and also differs from the Step 0
fixture theme (``com_helpers.apply_base_master``: accent1 C0504D,
Georgia/Arial, title 40pt bold centered, body 19/16/13, bullets en-dash/
square/guillemet) so profile-builder tests can tell the two corpora
apart and an Office-default fallback can never accidentally pass:

* palette: teal/amber scheme (accent1 1B7F79, not 4472C4 / C0504D)
* fonts: Georgia major / Calibri minor (fixtures use Arial minor)
* type scale {11, 14, 20, 30}pt -- no member equals the 18pt default
* body space_after 8pt (not 0 / not 10), line spacing 1.20
* bullets l1 em dash / l2 middle dot / l3 ">" (all typable, non-default)
* borders 1.25pt dashed #14324F, corner radius adj 0.12
* 3-column grid: lefts 60/360/660pt, rights 300/600/900pt (col 240,
  gutter 60), everything snapped
"""

from __future__ import annotations

import os

import _bootstrap  # noqa: F401  (sys.path side effects)
from com_helpers import (
    ALIGN_LEFT,
    ALIGN_RIGHT,
    LINE_DASH,
    MSO_FALSE,
    MSO_THEME_COLOR_ACCENT1,
    MSO_THEME_COLOR_ACCENT2,
    MSO_THEME_COLOR_ACCENT3,
    MSO_THEME_COLOR_ACCENT4,
    MSO_THEME_COLOR_ACCENT5,
    MSO_THEME_COLOR_ACCENT6,
    MSO_THEME_COLOR_DARK1,
    MSO_THEME_COLOR_DARK2,
    MSO_TRUE,
    PH_BODY,
    PH_OBJECT,
    SHAPE_RECTANGLE,
    SHAPE_ROUNDED_RECTANGLE,
    THEME_ACCENT1,
    THEME_ACCENT2,
    THEME_ACCENT3,
    THEME_ACCENT4,
    THEME_ACCENT5,
    THEME_ACCENT6,
    THEME_DARK1,
    THEME_DARK2,
    THEME_FOLLOWED_HYPERLINK,
    THEME_HYPERLINK,
    THEME_LIGHT1,
    THEME_LIGHT2,
    add_autoshape,
    add_textbox,
    apply_color_scheme,
    apply_font_scheme,
    find_placeholder,
    paragraph,
    paragraphs_count,
    set_adjustment,
    set_font,
    set_paragraph,
    set_title,
    solid_fill,
    style_line,
    style_master_body_level,
    style_master_title,
    no_line,
)

# ---------------------------------------------------------------------------
# Theme (custom color scheme + font scheme)
# ---------------------------------------------------------------------------

HOUSE_MASTER_NAME = "MeridianHouse"

#: Scheme token -> RRGGBB (no '#'). Serialized into corpus_truth.json.
SCHEME_HEX: dict[str, str] = {
    "dk1": "20262B",
    "lt1": "FAF9F6",
    "dk2": "14324F",
    "lt2": "DCE3E8",
    "accent1": "1B7F79",
    "accent2": "D97C2B",
    "accent3": "3E5C76",
    "accent4": "8A5A83",
    "accent5": "5E8C61",
    "accent6": "B23A48",
    "hlink": "176B87",
    "folHlink": "6D5D8F",
}

HOUSE_COLOR_SCHEME: dict[int, str] = {
    THEME_DARK1: SCHEME_HEX["dk1"],
    THEME_LIGHT1: SCHEME_HEX["lt1"],
    THEME_DARK2: SCHEME_HEX["dk2"],
    THEME_LIGHT2: SCHEME_HEX["lt2"],
    THEME_ACCENT1: SCHEME_HEX["accent1"],
    THEME_ACCENT2: SCHEME_HEX["accent2"],
    THEME_ACCENT3: SCHEME_HEX["accent3"],
    THEME_ACCENT4: SCHEME_HEX["accent4"],
    THEME_ACCENT5: SCHEME_HEX["accent5"],
    THEME_ACCENT6: SCHEME_HEX["accent6"],
    THEME_HYPERLINK: SCHEME_HEX["hlink"],
    THEME_FOLLOWED_HYPERLINK: SCHEME_HEX["folHlink"],
}

HOUSE_MAJOR_FONT = "Georgia"   # titles
HOUSE_MINOR_FONT = "Calibri"   # body / captions / footers

#: Scheme token -> MsoThemeColorIndex for ObjectThemeColor writes.
THEME_TOKEN_TO_MSO: dict[str, int] = {
    "dk1": MSO_THEME_COLOR_DARK1,
    "dk2": MSO_THEME_COLOR_DARK2,
    "accent1": MSO_THEME_COLOR_ACCENT1,
    "accent2": MSO_THEME_COLOR_ACCENT2,
    "accent3": MSO_THEME_COLOR_ACCENT3,
    "accent4": MSO_THEME_COLOR_ACCENT4,
    "accent5": MSO_THEME_COLOR_ACCENT5,
    "accent6": MSO_THEME_COLOR_ACCENT6,
}

# ---------------------------------------------------------------------------
# Type scale + master text styles (none of these equal Office defaults)
# ---------------------------------------------------------------------------

TYPE_SCALE_PT = (11.0, 14.0, 20.0, 30.0)

TITLE_SIZE_PT = 30.0        # master titleStyle (default is 44)
SUBTITLE_SIZE_PT = 20.0
BODY_SIZE_PT = 14.0         # master bodyStyle level 1 (default is 28)
CAPTION_SIZE_PT = 11.0      # footers, kickers, exhibit captions

TITLE_BOLD = False
TITLE_ALIGNMENT = "left"    # default master title alignment is centered
TITLE_COLOR_TOKEN = "dk2"

#: Master bodyStyle seeds per indent level (all non-default; space_after
#: is neither 0 nor 10; bullets are typable non-default characters).
BODY_LEVELS: dict[int, dict] = {
    1: {"size_pt": 14.0, "color_token": "dk1", "space_before_pt": 2.0,
        "space_after_pt": 8.0, "line_spacing": 1.20,
        "bullet_char": 0x2014, "bullet_glyph": "—",  # em dash
        "bullet_font": "Arial", "bullet_rel_size": 0.95},
    2: {"size_pt": 11.0, "color_token": "dk1", "space_before_pt": 2.0,
        "space_after_pt": 5.0, "line_spacing": 1.20,
        "bullet_char": 0x00B7, "bullet_glyph": "·",  # middle dot
        "bullet_font": "Arial", "bullet_rel_size": 0.90},
    3: {"size_pt": 11.0, "color_token": "dk1", "space_before_pt": 1.0,
        "space_after_pt": 4.0, "line_spacing": 1.20,
        "bullet_char": 0x003E, "bullet_glyph": ">",       # greater-than
        "bullet_font": "Arial", "bullet_rel_size": 0.90},
}

# ---------------------------------------------------------------------------
# Shape defaults
# ---------------------------------------------------------------------------

BORDER_WEIGHT_PT = 1.25
BORDER_COLOR_HEX = SCHEME_HEX["dk2"]     # 14324F
BORDER_DASH = LINE_DASH                  # OOXML "dash" (default is solid)
BORDER_DASH_NAME = "dash"
CORNER_RADIUS_ADJ = 0.12                 # roundRect adj (default 0.16667)
PANEL_FILL_HEX = SCHEME_HEX["lt2"]       # DCE3E8

# ---------------------------------------------------------------------------
# Alignment grid (points; slide is 960 x 540 pt, 16:9)
# ---------------------------------------------------------------------------

SLIDE_W_PT = 960.0
SLIDE_H_PT = 540.0

GRID_LEFT_EDGES_PT = (60.0, 360.0, 660.0)
GRID_RIGHT_EDGES_PT = (300.0, 600.0, 900.0)
GRID_CENTER_EDGES_PT = (180.0, 480.0, 780.0)
GRID_COLUMN_W_PT = 240.0
GRID_GUTTER_PT = 60.0
GRID_TOLERANCE_PT = 4.0

# ---------------------------------------------------------------------------
# Archetype geometry (x, y, w, h in points -- all edges grid-snapped)
# ---------------------------------------------------------------------------

TITLE_BAND = (60.0, 30.0, 840.0, 60.0)          # agenda/content/two_column
TITLE_MAIN = (60.0, 180.0, 840.0, 70.0)         # title archetype
TITLE_SUBTITLE = (60.0, 270.0, 540.0, 40.0)
TITLE_RULE = (60.0, 150.0, 240.0, 6.0)
DIVIDER_KICKER = (60.0, 200.0, 240.0, 24.0)
DIVIDER_TITLE = (60.0, 240.0, 840.0, 60.0)
DIVIDER_RULE = (60.0, 320.0, 840.0, 6.0)
BODY_FULL = (60.0, 120.0, 840.0, 240.0)         # content archetype body
BODY_AGENDA = (60.0, 120.0, 540.0, 330.0)       # agenda list (cols 1-2)
TAKEAWAY_PANEL = (60.0, 380.0, 840.0, 70.0)
COLUMN_PANEL_LEFT = (60.0, 120.0, 240.0, 300.0)
COLUMN_PANEL_RIGHT = (360.0, 120.0, 240.0, 300.0)
SIDEBAR_IMAGE = (660.0, 120.0, 240.0, 180.0)
SIDEBAR_CAPTION = (660.0, 315.0, 240.0, 24.0)
CLOSING_RULE = (60.0, 180.0, 240.0, 6.0)
CLOSING_TITLE = (60.0, 210.0, 840.0, 60.0)
CONTACT_PANEL = (360.0, 290.0, 240.0, 140.0)
FOOTER_ZONE = (60.0, 500.0, 840.0, 28.0)
FOOTER_NOTE = (60.0, 502.0, 300.0, 22.0)
FOOTER_PAGE = (780.0, 502.0, 120.0, 22.0)

FOOTER_NOTE_TEXT = "Meridian Advisory | synthetic corpus"
FOOTER_ARCHETYPES = ("agenda", "content", "two_column", "closing")

ARCHETYPES = ("title", "agenda", "section_divider", "content",
              "two_column", "closing")

# ---------------------------------------------------------------------------
# Master application
# ---------------------------------------------------------------------------

def apply_house_master(master) -> None:
    """Apply the Meridian theme + master text styles + placeholder bands."""
    master.Name = HOUSE_MASTER_NAME
    apply_color_scheme(master, HOUSE_COLOR_SCHEME)
    apply_font_scheme(master, HOUSE_MAJOR_FONT, HOUSE_MINOR_FONT)

    style_master_title(
        master,
        size=TITLE_SIZE_PT,
        bold=TITLE_BOLD,
        alignment=ALIGN_LEFT,
        theme_color=THEME_TOKEN_TO_MSO[TITLE_COLOR_TOKEN],
    )
    for level_index, seed in BODY_LEVELS.items():
        style_master_body_level(
            master, level_index,
            size=seed["size_pt"],
            theme_color=THEME_TOKEN_TO_MSO[seed["color_token"]],
            space_before_pt=seed["space_before_pt"],
            space_after_pt=seed["space_after_pt"],
            space_within_multiple=seed["line_spacing"],
            bullet_char=seed["bullet_char"],
            bullet_font=seed["bullet_font"],
            bullet_rel_size=seed["bullet_rel_size"],
        )
    _seed_master_bands(master)


def _seed_master_bands(master) -> None:
    """Move the master's title/body placeholders onto the house bands."""
    title = master.Shapes.Title
    title.Left, title.Top, title.Width, title.Height = TITLE_BAND
    body = find_placeholder(master, PH_BODY, PH_OBJECT)
    if body is None:
        raise LookupError(
            f"Master {master.Name!r} has no body placeholder to seed")
    body.Left, body.Top, body.Width, body.Height = BODY_FULL


# ---------------------------------------------------------------------------
# Shared slide-element builders (all snap to the grid rects above)
# ---------------------------------------------------------------------------

def place_title(slide, text: str, rect: tuple = TITLE_BAND):
    """Set the slide title text and snap the placeholder to a band."""
    set_title(slide, text)
    title = slide.Shapes.Title
    title.Left, title.Top, title.Width, title.Height = rect
    return title


def place_body(slide, rect: tuple):
    """Snap the body placeholder to a region; returns the placeholder."""
    body = find_placeholder(slide, PH_BODY, PH_OBJECT)
    if body is None:
        raise LookupError("Slide has no body placeholder")
    body.Name = "BodyContent"
    body.Left, body.Top, body.Width, body.Height = rect
    return body


def add_panel(slide, rect: tuple, name: str, header: str,
              lines: list[str], fill_hex: str = PANEL_FILL_HEX):
    """House content panel: rounded rect, lt2 fill, 1.25pt dashed border,
    corner adj 0.12, 14pt dk1 text, bold header paragraph, left aligned."""
    shape = add_autoshape(slide, SHAPE_ROUNDED_RECTANGLE, *rect, name=name)
    set_adjustment(shape, 1, CORNER_RADIUS_ADJ)
    solid_fill(shape, fill_hex)
    style_line(shape, weight_pt=BORDER_WEIGHT_PT, dash_style=BORDER_DASH,
               color_hex=BORDER_COLOR_HEX)
    text_range = shape.TextFrame2.TextRange
    text_range.Text = "\r".join([header] + list(lines))
    for i in range(1, paragraphs_count(text_range) + 1):
        para = paragraph(text_range, i)
        set_paragraph(para, alignment=ALIGN_LEFT)
        set_font(para, size=BODY_SIZE_PT, bold=(i == 1),
                 theme_color=MSO_THEME_COLOR_DARK1)
    return shape


def add_rule(slide, rect: tuple, name: str = "AccentRule"):
    """Thin accent1-filled rectangle used as a horizontal rule."""
    shape = add_autoshape(slide, SHAPE_RECTANGLE, *rect, name=name)
    solid_fill(shape, theme_color=MSO_THEME_COLOR_ACCENT1)
    no_line(shape)
    return shape


def add_kicker(slide, text: str, rect: tuple = DIVIDER_KICKER,
               name: str = "SectionKicker"):
    """Small bold accent1 label above a divider title (11pt Calibri)."""
    box = add_textbox(slide, *rect, text, name=name)
    set_font(box.TextFrame2.TextRange, size=CAPTION_SIZE_PT, bold=True,
             theme_color=MSO_THEME_COLOR_ACCENT1)
    return box


def add_caption(slide, text: str, rect: tuple = SIDEBAR_CAPTION,
                name: str = "SidebarCaption"):
    """Exhibit caption: 11pt italic accent3."""
    box = add_textbox(slide, *rect, text, name=name)
    set_font(box.TextFrame2.TextRange, size=CAPTION_SIZE_PT, italic=True,
             theme_color=MSO_THEME_COLOR_ACCENT3)
    return box


def add_subtitle(slide, text: str, rect: tuple = TITLE_SUBTITLE,
                 name: str = "SubtitleBox"):
    """Title-slide subtitle: 20pt accent3 Calibri."""
    box = add_textbox(slide, *rect, text, name=name)
    set_font(box.TextFrame2.TextRange, size=SUBTITLE_SIZE_PT,
             theme_color=MSO_THEME_COLOR_ACCENT3)
    return box


def add_footer(slide, page_no: int) -> None:
    """Footer zone convention: source note left, page number right."""
    note = add_textbox(slide, *FOOTER_NOTE, FOOTER_NOTE_TEXT,
                       name="FooterNote")
    set_font(note.TextFrame2.TextRange, size=CAPTION_SIZE_PT,
             theme_color=MSO_THEME_COLOR_ACCENT3)
    page = add_textbox(slide, *FOOTER_PAGE, f"{page_no:02d}",
                       name="FooterPage")
    page_range = page.TextFrame2.TextRange
    set_font(page_range, size=CAPTION_SIZE_PT,
             theme_color=MSO_THEME_COLOR_ACCENT3)
    set_paragraph(paragraph(page_range, 1), alignment=ALIGN_RIGHT)


def add_image(slide, png_path: str, rect: tuple = SIDEBAR_IMAGE,
              name: str = "SidebarExhibit"):
    """Embed a generated PNG asset into an image zone."""
    path = os.path.abspath(png_path)
    if not os.path.exists(path):
        raise FileNotFoundError(f"image asset missing: {path}")
    pic = slide.Shapes.AddPicture(path, MSO_FALSE, MSO_TRUE, *rect)
    pic.Name = name
    return pic


# ---------------------------------------------------------------------------
# Theme-color tint helpers (PowerPoint's "Lighter/Darker N%" variants).
# Brightness b > 0  ->  lumMod (1-b) + lumOff b;  b < 0  ->  lumMod (1+b).
# ---------------------------------------------------------------------------

def tint_fill(shape, token: str, brightness: float) -> None:
    """Solid-fill a shape with a theme color variant (writes lumMod/lumOff)."""
    solid_fill(shape, theme_color=THEME_TOKEN_TO_MSO[token])
    shape.Fill.ForeColor.Brightness = brightness


def tint_font(text_range, token: str, brightness: float) -> None:
    """Color a text range with a theme color variant."""
    fore = text_range.Font.Fill.ForeColor
    fore.ObjectThemeColor = THEME_TOKEN_TO_MSO[token]
    fore.Brightness = brightness


def brightness_transforms(brightness: float) -> dict[str, int]:
    """Brightness -> the lumMod/lumOff vals PowerPoint writes (probe-verified)."""
    if brightness > 0:
        return {"lumMod": round((1.0 - brightness) * 100000),
                "lumOff": round(brightness * 100000)}
    return {"lumMod": round((1.0 + brightness) * 100000)}
