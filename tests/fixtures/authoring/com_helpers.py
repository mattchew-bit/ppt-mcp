"""Shared COM helpers for authoring the style-fidelity fixture decks.

These utilities drive *desktop PowerPoint* via COM (pywin32) so that the
fixture .pptx files are written by PowerPoint itself.  The Step 0 plan
forbids python-pptx-generated fixtures: real inheritance structures
(theme -> master txStyles -> layout -> slide) must be produced by the
application whose behavior the effective-style resolver is later verified
against.

COM hygiene enforced here:
- ``pythoncom.CoInitialize()`` before any COM call, paired CoUninitialize.
- Never sets ``app.Visible = False`` (PowerPoint raises).
- Operates only on Presentation objects created/opened by these helpers.
- ``app.Quit()`` only when ``Presentations.Count == 0`` after our work, so
  a user's open decks are never disturbed.
- ``AutomationSecurity = msoAutomationSecurityForceDisable`` while working.
- Absolute paths everywhere (COM resolves relative paths against the
  PowerPoint process CWD, not Python's).
"""

from __future__ import annotations

import contextlib
import os

import pythoncom
import win32com.client
from win32com.client import gencache

# --------------------------------------------------------------------------
# Office / PowerPoint enum constants (raw ints -- avoids makepy dependency)
# --------------------------------------------------------------------------

MSO_TRUE = -1
MSO_FALSE = 0

PP_SAVE_AS_OPENXML = 24          # ppSaveAsOpenXMLPresentation
PP_ALERTS_NONE = 1               # ppAlertsNone
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# MsoThemeColorSchemeIndex (Theme.ThemeColorScheme.Colors(index))
THEME_DARK1 = 1
THEME_LIGHT1 = 2
THEME_DARK2 = 3
THEME_LIGHT2 = 4
THEME_ACCENT1 = 5
THEME_ACCENT2 = 6
THEME_ACCENT3 = 7
THEME_ACCENT4 = 8
THEME_ACCENT5 = 9
THEME_ACCENT6 = 10
THEME_HYPERLINK = 11
THEME_FOLLOWED_HYPERLINK = 12

# MsoThemeColorIndex (ColorFormat.ObjectThemeColor)
MSO_THEME_COLOR_DARK1 = 1
MSO_THEME_COLOR_LIGHT1 = 2
MSO_THEME_COLOR_DARK2 = 3
MSO_THEME_COLOR_LIGHT2 = 4
MSO_THEME_COLOR_ACCENT1 = 5
MSO_THEME_COLOR_ACCENT2 = 6
MSO_THEME_COLOR_ACCENT3 = 7
MSO_THEME_COLOR_ACCENT4 = 8
MSO_THEME_COLOR_ACCENT5 = 9
MSO_THEME_COLOR_ACCENT6 = 10

MSO_THEME_LATIN = 1              # ThemeFonts.Item index for the latin script

# PpTextStyleType (SlideMaster.TextStyles)
PP_DEFAULT_STYLE = 1
PP_TITLE_STYLE = 2
PP_BODY_STYLE = 3

# PpPlaceholderType
PH_TITLE = 1
PH_BODY = 2
PH_CENTER_TITLE = 3
PH_SUBTITLE = 4
PH_OBJECT = 7

# MsoShapeType
SHAPE_TYPE_AUTOSHAPE = 1
SHAPE_TYPE_PLACEHOLDER = 14
SHAPE_TYPE_TEXTBOX = 17

# MsoAutoShapeType
SHAPE_RECTANGLE = 1
SHAPE_ROUNDED_RECTANGLE = 5
SHAPE_OVAL = 9
SHAPE_CHEVRON = 52

MSO_TEXT_ORIENTATION_HORIZONTAL = 1

# MsoLineDashStyle
LINE_SOLID = 1
LINE_SQUARE_DOT = 2
LINE_ROUND_DOT = 3
LINE_DASH = 4
LINE_DASH_DOT = 5
LINE_DASH_DOT_DOT = 6
LINE_LONG_DASH = 7
LINE_LONG_DASH_DOT = 8

# Paragraph alignment (PpParagraphAlignment / MsoParagraphAlignment, 1..4 agree)
ALIGN_LEFT = 1
ALIGN_CENTER = 2
ALIGN_RIGHT = 3
ALIGN_JUSTIFY = 4

# MsoGradientStyle
GRADIENT_HORIZONTAL = 1

# Bullet characters used across the fixtures (unicode escapes, Arial-safe)
BULLET_EN_DASH = 0x2013          # "-" en dash
BULLET_FILLED_SQUARE = 0x25A0    # black square
BULLET_RIGHT_GUILLEMET = 0x00BB  # right-pointing double angle
BULLET_BLACK_DIAMOND = 0x25C6
BULLET_SMALL_SQUARE = 0x25AA
BULLET_MULTIPLY = 0x00D7
BULLET_WHITE_CIRCLE = 0x25CB
BULLET_WHITE_BULLET = 0x25E6
BULLET_CHECK_MARK = 0x2713


# --------------------------------------------------------------------------
# Color conversion
# --------------------------------------------------------------------------

def ole_rgb(hex_color: str) -> int:
    """'1F4E79' -> OLE color long (0x00BBGGRR byte order)."""
    value = hex_color.lstrip("#")
    if len(value) != 6:
        raise ValueError(f"Expected RRGGBB hex color, got {hex_color!r}")
    r = int(value[0:2], 16)
    g = int(value[2:4], 16)
    b = int(value[4:6], 16)
    return r | (g << 8) | (b << 16)


def hex_from_ole(value: int) -> str:
    """OLE color long -> 'RRGGBB'."""
    r = value & 0xFF
    g = (value >> 8) & 0xFF
    b = (value >> 16) & 0xFF
    return f"{r:02X}{g:02X}{b:02X}"


# --------------------------------------------------------------------------
# Parameterized COM property access
#
# TextRange2.Paragraphs / Runs / Characters are parameterized *properties*
# with optional arguments.  makepy places those in _prop_map_get_, so plain
# attribute access invokes them with default args and the index can never
# be passed; dynamic dispatch fails outright ("Does not support a
# collection").  The only reliable route is a low-level Invoke.
# --------------------------------------------------------------------------

def com_get(obj, name: str, *args):
    """Invoke a (possibly parameterized) COM property get / method."""
    ole = getattr(obj, "_oleobj_", obj)
    dispid = ole.GetIDsOfNames(0, name)
    flags = pythoncom.DISPATCH_METHOD | pythoncom.DISPATCH_PROPERTYGET
    result = ole.Invoke(dispid, 0, flags, 1, *args)
    if isinstance(result, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]):
        return win32com.client.Dispatch(result)
    return result


def paragraph(text_range, index: int):
    """TextRange2.Paragraphs(index) -- a single paragraph range."""
    return com_get(text_range, "Paragraphs", index)


def paragraphs_count(text_range) -> int:
    return int(com_get(text_range, "Paragraphs").Count)


def run_item(paragraph_range, index: int):
    """TextRange2.Runs(index, 1) -- a single run range."""
    return com_get(paragraph_range, "Runs", index, 1)


def runs_count(paragraph_range) -> int:
    return int(com_get(paragraph_range, "Runs").Count)


def characters(text_range, start: int, length: int):
    """TextRange2.Characters(start, length) -- 1-based character subrange."""
    return com_get(text_range, "Characters", start, length)


# --------------------------------------------------------------------------
# Application / presentation lifecycle
# --------------------------------------------------------------------------

# The Office (MSO) type library hosts TextFrame2/TextRange2/Font2 etc.
# Its makepy module must exist as well, otherwise those objects fall back
# to dynamic dispatch, which cannot invoke parameterized properties such
# as TextRange2.Paragraphs(i) ("Does not support a collection").
_MSO_TYPELIB_GUID = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"


def _ensure_office_typelib() -> None:
    for minor in range(9, -1, -1):
        with contextlib.suppress(Exception):
            if gencache.EnsureModule(_MSO_TYPELIB_GUID, 0, 2, minor) is not None:
                return
    raise RuntimeError(
        "Could not generate makepy wrappers for the Office (MSO) type "
        "library; TextRange2 access would fail under dynamic dispatch."
    )


@contextlib.contextmanager
def powerpoint_app():
    """Yield a PowerPoint.Application with alerts off and macros disabled.

    Never toggles ``Visible`` (setting it False throws).  Quits the
    application afterwards only if no presentations remain open, so a
    running PowerPoint session belonging to the user is left untouched.
    """
    pythoncom.CoInitialize()
    # EnsureDispatch (makepy static wrappers) rather than plain Dispatch:
    # parameterized properties are only callable through generated wrappers.
    _ensure_office_typelib()
    app = gencache.EnsureDispatch("PowerPoint.Application")
    saved_alerts = None
    saved_security = None
    try:
        with contextlib.suppress(Exception):
            saved_alerts = app.DisplayAlerts
            app.DisplayAlerts = PP_ALERTS_NONE
        with contextlib.suppress(Exception):
            saved_security = app.AutomationSecurity
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        yield app
    finally:
        with contextlib.suppress(Exception):
            if saved_alerts is not None:
                app.DisplayAlerts = saved_alerts
        with contextlib.suppress(Exception):
            if saved_security is not None:
                app.AutomationSecurity = saved_security
        with contextlib.suppress(Exception):
            if app.Presentations.Count == 0:
                app.Quit()
        del app
        pythoncom.CoUninitialize()


@contextlib.contextmanager
def new_presentation(app):
    """Create a fresh presentation (no window) and always close it."""
    pres = app.Presentations.Add(MSO_FALSE)
    try:
        yield pres
    finally:
        with contextlib.suppress(Exception):
            pres.Saved = MSO_TRUE  # suppress save prompt on close
        with contextlib.suppress(Exception):
            pres.Close()


@contextlib.contextmanager
def open_presentation(app, path: str, read_only: bool = True):
    """Open an existing deck (no window, read-only by default), always close."""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(abs_path)
    pres = app.Presentations.Open(
        abs_path,
        MSO_TRUE if read_only else MSO_FALSE,  # ReadOnly
        MSO_FALSE,                             # Untitled
        MSO_FALSE,                             # WithWindow
    )
    try:
        yield pres
    finally:
        with contextlib.suppress(Exception):
            pres.Saved = MSO_TRUE
        with contextlib.suppress(Exception):
            pres.Close()


def save_pptx(pres, path: str) -> str:
    """SaveAs .pptx (format 24), overwriting any previous output (idempotent)."""
    abs_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    if os.path.exists(abs_path):
        try:
            os.remove(abs_path)
        except PermissionError as exc:
            raise PermissionError(
                f"Cannot overwrite {abs_path}: file is locked "
                "(is it open in PowerPoint?)"
            ) from exc
    pres.SaveAs(abs_path, PP_SAVE_AS_OPENXML)
    return abs_path


def run_standalone(build_fn, *args, **kwargs):
    """Run a build(app, ...) function inside its own application context."""
    with powerpoint_app() as app:
        return build_fn(app, *args, **kwargs)


# --------------------------------------------------------------------------
# Theme / master customization
# --------------------------------------------------------------------------

def apply_color_scheme(master, colors: dict[int, str]) -> None:
    scheme = master.Theme.ThemeColorScheme
    for index, hex_color in colors.items():
        scheme.Colors(index).RGB = ole_rgb(hex_color)


def apply_font_scheme(master, major_latin: str, minor_latin: str) -> None:
    font_scheme = master.Theme.ThemeFontScheme
    font_scheme.MajorFont.Item(MSO_THEME_LATIN).Name = major_latin
    font_scheme.MinorFont.Item(MSO_THEME_LATIN).Name = minor_latin


def _apply_color(color_format, color_hex: str | None, theme_color: int | None):
    if theme_color is not None:
        color_format.ObjectThemeColor = theme_color
    elif color_hex is not None:
        color_format.RGB = ole_rgb(color_hex)


def style_master_title(
    master,
    *,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    alignment: int | None = None,
    color_hex: str | None = None,
    theme_color: int | None = None,
) -> None:
    """Seed the master title text style (writes txStyles/titleStyle)."""
    level = master.TextStyles(PP_TITLE_STYLE).Levels(1)
    if size is not None:
        level.Font.Size = size
    if bold is not None:
        level.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    if italic is not None:
        level.Font.Italic = MSO_TRUE if italic else MSO_FALSE
    _apply_color(level.Font.Color, color_hex, theme_color)
    if alignment is not None:
        level.ParagraphFormat.Alignment = alignment


def style_master_body_level(
    master,
    level_index: int,
    *,
    size: float | None = None,
    color_hex: str | None = None,
    theme_color: int | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    space_within_multiple: float | None = None,
    bullet_char: int | None = None,
    bullet_font: str = "Arial",
    bullet_rel_size: float | None = None,
    alignment: int | None = None,
) -> None:
    """Seed one level of the master body text style (writes txStyles/bodyStyle).

    Spacing values are POINTS: the line-rule flags are forced off before the
    values are written (PowerPoint interprets Space* as line counts when the
    matching LineRule* flag is msoTrue).  Line spacing is a MULTIPLE
    (LineRuleWithin stays msoTrue).
    """
    level = master.TextStyles(PP_BODY_STYLE).Levels(level_index)
    font = level.Font
    para = level.ParagraphFormat

    if size is not None:
        font.Size = size
    _apply_color(font.Color, color_hex, theme_color)

    if alignment is not None:
        para.Alignment = alignment
    if space_before_pt is not None:
        para.LineRuleBefore = MSO_FALSE
        para.SpaceBefore = space_before_pt
    if space_after_pt is not None:
        para.LineRuleAfter = MSO_FALSE
        para.SpaceAfter = space_after_pt
    if space_within_multiple is not None:
        para.LineRuleWithin = MSO_TRUE
        para.SpaceWithin = space_within_multiple

    if bullet_char is not None:
        bullet = para.Bullet
        bullet.Visible = MSO_TRUE
        bullet.Character = bullet_char
        bullet.Font.Name = bullet_font
        if bullet_rel_size is not None:
            bullet.RelativeSize = bullet_rel_size


# --------------------------------------------------------------------------
# Base fixture theme -- shared by theme_only / layout_override /
# explicit_override and by master #1 of multi_master.
#
# NULL-TEST GUARD: every value below deliberately differs from the Office
# defaults (Office theme: Calibri Light/Calibri, accent1 4472C4, title 44pt,
# body 28/24/20pt, spcBef 10pt, bullet "•") so a resolver that returns
# defaults can never accidentally pass.
# --------------------------------------------------------------------------

BASE_COLOR_SCHEME: dict[int, str] = {
    THEME_DARK1: "1A1A2E",
    THEME_LIGHT1: "FDFBF7",
    THEME_DARK2: "1F4E79",
    THEME_LIGHT2: "E8E4D8",
    THEME_ACCENT1: "C0504D",
    THEME_ACCENT2: "6B9F59",
    THEME_ACCENT3: "7C5CA6",
    THEME_ACCENT4: "E2A33D",
    THEME_ACCENT5: "3E8FB0",
    THEME_ACCENT6: "A63D57",
    THEME_HYPERLINK: "2E86AB",
    THEME_FOLLOWED_HYPERLINK: "8C5E93",
}

BASE_MAJOR_FONT = "Georgia"
BASE_MINOR_FONT = "Arial"


def apply_base_master(master) -> None:
    """Apply the shared fixture theme + master text styles to a slide master."""
    apply_color_scheme(master, BASE_COLOR_SCHEME)
    apply_font_scheme(master, BASE_MAJOR_FONT, BASE_MINOR_FONT)

    # Title: 40pt (not 44), bold (not regular), centered (not left),
    # accent1 scheme color (resolves to C0504D, not dk1).
    style_master_title(
        master,
        size=40,
        bold=True,
        alignment=ALIGN_CENTER,
        theme_color=MSO_THEME_COLOR_ACCENT1,
    )

    # Body levels: sizes 19/16/13 (not 28/24/20), spacing 5/9, 4/7, 3/5 pt
    # (not 10/0), line spacing 1.15/1.10/1.05 (not 0.9), bullets en-dash /
    # filled-square / guillemet (not "•").
    style_master_body_level(
        master, 1,
        size=19,
        theme_color=MSO_THEME_COLOR_DARK2,   # resolves to 1F4E79
        space_before_pt=5, space_after_pt=9, space_within_multiple=1.15,
        bullet_char=BULLET_EN_DASH, bullet_rel_size=0.9,
    )
    style_master_body_level(
        master, 2,
        size=16,
        color_hex="3F3F66",
        space_before_pt=4, space_after_pt=7, space_within_multiple=1.10,
        bullet_char=BULLET_FILLED_SQUARE, bullet_rel_size=0.8,
    )
    style_master_body_level(
        master, 3,
        size=13,
        space_before_pt=3, space_after_pt=5, space_within_multiple=1.05,
        bullet_char=BULLET_RIGHT_GUILLEMET, bullet_rel_size=1.0,
    )


# --------------------------------------------------------------------------
# Layout / slide content helpers
# --------------------------------------------------------------------------

def find_layout(master, name_fragment: str):
    """Return the first custom layout whose name contains name_fragment."""
    fragment = name_fragment.lower()
    names = []
    for i in range(1, master.CustomLayouts.Count + 1):
        layout = master.CustomLayouts(i)
        names.append(layout.Name)
        if fragment in layout.Name.lower():
            return layout
    raise LookupError(
        f"No layout matching {name_fragment!r} on master {master.Name!r}; "
        f"available: {names}"
    )


def find_placeholder(shapes_owner, *ph_types: int):
    """Return the first placeholder of any given type on a slide or layout."""
    placeholders = shapes_owner.Shapes.Placeholders
    for i in range(1, placeholders.Count + 1):
        ph = placeholders(i)
        if ph.PlaceholderFormat.Type in ph_types:
            return ph
    return None


def add_slide(pres, layout):
    """Append a slide bound to the given custom layout."""
    return pres.Slides.AddSlide(pres.Slides.Count + 1, layout)


def set_title(slide, text: str) -> None:
    slide.Shapes.Title.TextFrame2.TextRange.Text = text


def set_bullet_text(shape, items: list[tuple[int, str]]) -> None:
    """Fill a body placeholder with (indent_level, text) paragraphs.

    Only structural properties (text + indent level) are written; all
    styling is inherited.
    """
    text_range = shape.TextFrame2.TextRange
    text_range.Text = "\r".join(text for _, text in items)
    for i, (level, _) in enumerate(items, start=1):
        paragraph(text_range, i).ParagraphFormat.IndentLevel = level


def add_textbox(slide, left: float, top: float, width: float, height: float,
                text: str, name: str | None = None):
    box = slide.Shapes.AddTextbox(
        MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)
    box.TextFrame2.TextRange.Text = text
    if name:
        box.Name = name
    return box


def add_autoshape(slide, shape_type: int, left: float, top: float,
                  width: float, height: float, text: str | None = None,
                  name: str | None = None):
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
    if text is not None:
        shape.TextFrame2.TextRange.Text = text
    if name:
        shape.Name = name
    return shape


def set_adjustment(shape, index: int, value: float) -> None:
    """Set a shape geometry adjustment (e.g. rounded-corner radius).

    ``Adjustments.Item`` is a parameterized property put, which neither
    dynamic nor static pywin32 wrappers expose as plain attribute
    assignment -- invoke the default member (DISPID 0) directly.
    """
    adjustments = shape.Adjustments
    ole = getattr(adjustments, "_oleobj_", adjustments)
    ole.Invoke(0, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, index, float(value))


def style_line(shape, *, weight_pt: float, dash_style: int = LINE_SOLID,
               color_hex: str | None = None, theme_color: int | None = None) -> None:
    line = shape.Line
    line.Visible = MSO_TRUE
    line.Weight = weight_pt
    line.DashStyle = dash_style
    _apply_color(line.ForeColor, color_hex, theme_color)


def no_line(shape) -> None:
    shape.Line.Visible = MSO_FALSE


def solid_fill(shape, color_hex: str | None = None,
               theme_color: int | None = None) -> None:
    fill = shape.Fill
    fill.Visible = MSO_TRUE
    fill.Solid()
    _apply_color(fill.ForeColor, color_hex, theme_color)


def gradient_fill(shape, base_color_hex: str, degree: float = 0.6) -> None:
    fill = shape.Fill
    fill.Visible = MSO_TRUE
    fill.ForeColor.RGB = ole_rgb(base_color_hex)
    fill.OneColorGradient(GRADIENT_HORIZONTAL, 1, degree)


def set_font(text_range, *, name: str | None = None, size: float | None = None,
             bold: bool | None = None, italic: bool | None = None,
             color_hex: str | None = None, theme_color: int | None = None) -> None:
    """Explicit run-level formatting on a TextRange2."""
    font = text_range.Font
    if name is not None:
        font.Name = name
    if size is not None:
        font.Size = size
    if bold is not None:
        font.Bold = MSO_TRUE if bold else MSO_FALSE
    if italic is not None:
        font.Italic = MSO_TRUE if italic else MSO_FALSE
    if theme_color is not None:
        font.Fill.ForeColor.ObjectThemeColor = theme_color
    elif color_hex is not None:
        font.Fill.ForeColor.RGB = ole_rgb(color_hex)


def set_paragraph(paragraph_range, *, alignment: int | None = None,
                  space_before_pt: float | None = None,
                  space_after_pt: float | None = None,
                  space_within_multiple: float | None = None,
                  bullet_char: int | None = None,
                  bullet_font: str = "Arial",
                  bullet_rel_size: float | None = None,
                  bullet_color_hex: str | None = None,
                  bullet_visible: bool | None = None) -> None:
    """Explicit paragraph-level formatting on a TextRange2 paragraph."""
    pf = paragraph_range.ParagraphFormat
    if alignment is not None:
        pf.Alignment = alignment
    if space_before_pt is not None:
        pf.LineRuleBefore = MSO_FALSE
        pf.SpaceBefore = space_before_pt
    if space_after_pt is not None:
        pf.LineRuleAfter = MSO_FALSE
        pf.SpaceAfter = space_after_pt
    if space_within_multiple is not None:
        pf.LineRuleWithin = MSO_TRUE
        pf.SpaceWithin = space_within_multiple

    bullet = pf.Bullet
    if bullet_visible is not None:
        bullet.Visible = MSO_TRUE if bullet_visible else MSO_FALSE
    if bullet_char is not None:
        bullet.Visible = MSO_TRUE
        bullet.Character = bullet_char
        bullet.Font.Name = bullet_font
        if bullet_rel_size is not None:
            bullet.RelativeSize = bullet_rel_size
        if bullet_color_hex is not None:
            bullet.Font.Fill.ForeColor.RGB = ole_rgb(bullet_color_hex)
