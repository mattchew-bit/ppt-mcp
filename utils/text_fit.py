"""Standalone text-fit prediction (Step 5 ``predict_text_fit``).

Runs python-pptx's own ``TextFitter`` machinery STANDALONE -- no deck
mutation, unlike ``TextFrame.fit_text`` -- with the three known patches
(research 2026-07-03 §2):

(a) **font resolver**: python-pptx only scans the system font
    directory. This resolver additionally covers the per-user
    ``%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts`` directory and the
    ``HKLM ...\\Fonts`` registry, reads family/bold/italic straight
    from font name tables via fontTools (already a dependency), and
    falls back Arial -> Calibri -> common Linux faces so a verdict is
    still possible under substitution (substitutions are reported,
    never silent).
(b) **line height**: the stock fitter measures a rendered ``"Ty"``
    bounding box, which understates real line pitch. Lines are billed
    at ~1.2x the point size, honoring explicit ``lnSpc`` overrides
    (percent multiples scale the 1.2x pitch; point values are used
    as-is).
(c) **per-paragraph effective sizes**: each paragraph is wrapped and
    billed at its inheritance-RESOLVED size/font (Step 2 resolver
    facts), not one frame-wide size.

Verdict per frame: ``fits`` / ``overflow`` / ``borderline`` (required
height within ±5% of available -> "confirm by render") -- NEVER a hard
fail. Frames with ``spAutoFit`` report ``fits`` (PowerPoint grows the
shape, and the stored height is already the grown one); frames whose
fonts cannot be resolved at all report ``unknown``.

``.ttc`` collections are not indexed (Calibri/Arial/Georgia ship as
.ttf); faces inside collections resolve through the fallback chain.
"""

import os
import sys
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from pptx.oxml.ns import qn
from pptx.text.layout import TextFitter, _Line, _LineSource

from .lint_xml import autofit_of, wrap_of

#: Approximate single-spaced line pitch as a multiple of point size.
LINE_HEIGHT_FACTOR = 1.2

#: ±band around a full frame that yields the "borderline" verdict.
BORDERLINE_BAND = 0.05

#: Substitution chain when the requested family is not installed.
FALLBACK_FAMILIES = ("Arial", "Calibri", "DejaVu Sans",
                     "Liberation Sans")

#: Paragraph size used when a paragraph carries no resolvable runs.
DEFAULT_PARAGRAPH_SIZE_PT = 18.0

_EMU_PER_PT = 12700
_BODYPR_INSET_DEFAULTS_EMU = {"lIns": 91440, "tIns": 45720,
                              "rIns": 91440, "bIns": 45720}


# ---------------------------------------------------------------------------
# Patch (a): font resolver (dirs + registry, via fontTools)
# ---------------------------------------------------------------------------

def _windows_font_dirs() -> List[Path]:
    dirs = []
    windir = os.environ.get("WINDIR", r"C:\Windows")
    dirs.append(Path(windir) / "Fonts")
    local = os.environ.get("LOCALAPPDATA")
    if local:
        dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    return dirs


def _posix_font_dirs() -> List[Path]:
    return [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
            Path("/System/Library/Fonts"), Path("/Library/Fonts"),
            Path.home() / ".fonts"]


def _registry_font_files() -> Iterable[Path]:
    """Font files registered under HKLM (may live outside the dirs)."""
    if not sys.platform.startswith("win"):
        return []
    try:
        import winreg
    except ImportError:  # pragma: no cover - windows always has winreg
        return []
    files = []
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            index = 0
            while True:
                try:
                    _, value, _ = winreg.EnumValue(key, index)
                except OSError:
                    break
                index += 1
                if isinstance(value, str) and value:
                    path = Path(value)
                    if not path.is_absolute():
                        path = _windows_font_dirs()[0] / value
                    files.append(path)
    except OSError:
        return []
    return files


def _iter_font_files() -> Iterable[Path]:
    seen = set()
    dirs = (_windows_font_dirs() if sys.platform.startswith("win")
            else _posix_font_dirs())
    for directory in dirs:
        if not directory.is_dir():
            continue
        for path in sorted(directory.rglob("*")):
            if path.suffix.lower() in (".ttf", ".otf"):
                key = str(path).casefold()
                if key not in seen:
                    seen.add(key)
                    yield path
    for path in _registry_font_files():
        key = str(path).casefold()
        if (key not in seen and path.suffix.lower() in (".ttf", ".otf")
                and path.is_file()):
            seen.add(key)
            yield path


def _face_attributes(path: Path) -> Optional[Tuple[str, bool, bool]]:
    """(family, bold, italic) from a font file's name/head tables."""
    import logging

    from fontTools.ttLib import TTFont

    # Some installed fonts carry malformed head timestamps; fontTools
    # logs a warning per file, which is pure noise for indexing.
    logging.getLogger("fontTools.ttLib.tables._h_e_a_d").setLevel(
        logging.ERROR)
    try:
        font = TTFont(str(path), lazy=True, fontNumber=0)
        try:
            name = font["name"]
            family = (name.getDebugName(16) or name.getDebugName(1))
            mac_style = font["head"].macStyle
        finally:
            font.close()
    except Exception:
        return None
    if not family:
        return None
    return family, bool(mac_style & 1), bool(mac_style & 2)


@lru_cache(maxsize=1)
def _font_index() -> Dict:
    """{(family_cf, bold, italic): path} + {family_cf: regular path}."""
    styled: Dict[Tuple[str, bool, bool], str] = {}
    by_family: Dict[str, str] = {}
    regular_seen = set()
    for path in _iter_font_files():
        attrs = _face_attributes(path)
        if attrs is None:
            continue
        family, bold, italic = attrs
        family_cf = family.casefold()
        styled.setdefault((family_cf, bold, italic), str(path))
        if not bold and not italic:
            if family_cf not in regular_seen:
                regular_seen.add(family_cf)
                by_family[family_cf] = str(path)
        else:
            by_family.setdefault(family_cf, str(path))
    return {"styled": styled, "by_family": by_family}


def resolve_font_file(family: Optional[str], bold: bool = False,
                      italic: bool = False
                      ) -> Tuple[Optional[str], Optional[str]]:
    """(font_file_path, family_actually_used) for a requested face.

    Tries the exact style, then the family's regular face, then the
    fallback chain, then any indexed font at all (deterministic pick).
    Returns ``(None, None)`` only when no font file exists anywhere.
    """
    index = _font_index()
    candidates = [family] if family else []
    candidates.extend(f for f in FALLBACK_FAMILIES
                      if f.casefold() != (family or "").casefold())
    for candidate in candidates:
        family_cf = candidate.casefold()
        for style in ((family_cf, bold, italic), (family_cf, False,
                                                  False)):
            path = index["styled"].get(style)
            if path:
                return path, candidate
        path = index["by_family"].get(family_cf)
        if path:
            return path, candidate
    if index["by_family"]:
        family_cf = min(index["by_family"])
        return index["by_family"][family_cf], family_cf
    return None, None


# ---------------------------------------------------------------------------
# Patch (b): TextFitter subclass with real line pitch + break guard
# ---------------------------------------------------------------------------

class HouseTextFitter(TextFitter):
    """python-pptx ``TextFitter`` with the Step 5 line-height patch.

    ``line_pitch_pt`` (set after construction; tuple subclasses accept
    attributes) replaces the rendered-"Ty" height with the patched
    pitch. ``_break_line`` gains a guard so a single word wider than
    the frame becomes its own line instead of crashing the fitter.
    """

    line_pitch_factor: float = LINE_HEIGHT_FACTOR
    line_multiple: float = 1.0

    def _break_line(self, line_source, point_size):
        result = super()._break_line(line_source, point_size)
        if result is not None:
            return result
        words = line_source._text.split()
        return _Line(words[0], _LineSource(" ".join(words[1:])))

    @property
    def _fits_inside_predicate(self):
        def predicate(point_size):
            lines = self._wrap_lines(self._line_source, point_size)
            pitch_emu = int(point_size * self.line_pitch_factor
                            * self.line_multiple * _EMU_PER_PT)
            return (pitch_emu * len(lines)) <= self._height

        return predicate

    def wrap_line_count(self, text: str, point_size: int) -> int:
        source = _LineSource(text)
        if not source:
            return 1
        return len(self._wrap_lines(source, point_size))


# ---------------------------------------------------------------------------
# Patch (c): per-paragraph assessment over resolver facts
# ---------------------------------------------------------------------------

def _bodypr_insets_pt(shape_elem) -> Optional[Tuple[float, float]]:
    """(horizontal, vertical) inset totals in points, else None."""
    txbody = shape_elem.find(qn("p:txBody"))
    if txbody is None:
        return None
    bodypr = txbody.find(qn("a:bodyPr"))
    insets = {}
    for attr, default in _BODYPR_INSET_DEFAULTS_EMU.items():
        raw = bodypr.get(attr) if bodypr is not None else None
        insets[attr] = (int(raw) if raw is not None else default)
    horizontal = (insets["lIns"] + insets["rIns"]) / _EMU_PER_PT
    vertical = (insets["tIns"] + insets["bIns"]) / _EMU_PER_PT
    return horizontal, vertical


def _paragraph_face(paragraph: Dict, carry: Dict) -> Dict:
    """Effective face for a paragraph (first run; carry-over when
    runless)."""
    runs = paragraph.get("runs", [])
    if runs:
        font = runs[0]["font"]
        carry = {
            "size": max((run["font"].get("size_pt")
                         or DEFAULT_PARAGRAPH_SIZE_PT) for run in runs),
            "name": font.get("name"),
            "bold": bool(font.get("bold")),
            "italic": bool(font.get("italic")),
        }
    return carry


def _paragraph_height_pt(paragraph: Dict, lines: int, size: float) -> float:
    spacing = paragraph.get("line_spacing", {})
    if "points" in spacing:
        pitch = spacing["points"]
        multiple = 1.0
    else:
        multiple = spacing.get("multiple", 1.0)
        pitch = size * LINE_HEIGHT_FACTOR * multiple
    height = lines * pitch
    for key in ("space_before", "space_after"):
        entry = paragraph.get(key, {})
        if "points" in entry:
            height += entry["points"]
        elif "lines" in entry:
            height += entry["lines"] * pitch
    return height


def _frame_box_pt(record: Dict) -> Optional[Tuple[float, float]]:
    geometry = record.get("geometry", {})
    width, height = geometry.get("width_pt"), geometry.get("height_pt")
    if width is None or height is None:
        return None
    insets = _bodypr_insets_pt(record["_shape"]._element)
    if insets is None:
        return None
    avail_w, avail_h = width - insets[0], height - insets[1]
    if avail_w <= 0 or avail_h <= 0:
        return None
    return avail_w, avail_h


def _verdict(ratio: float) -> str:
    if ratio < 1.0 - BORDERLINE_BAND:
        return "fits"
    if ratio <= 1.0 + BORDERLINE_BAND:
        return "borderline"
    return "overflow"


def _required_height_pt(record: Dict, avail_w: float, avail_h: float,
                        wrap: Optional[str]
                        ) -> Tuple[Optional[float], List[str]]:
    """(total required height in points, substitutions) or (None, ...)
    when no font file is available to measure with."""
    required = 0.0
    substitutions: List[str] = []
    face: Dict = {"size": DEFAULT_PARAGRAPH_SIZE_PT, "name": None,
                  "bold": False, "italic": False}
    for paragraph in record.get("paragraphs", []):
        face = _paragraph_face(paragraph, face)
        font_file, used = resolve_font_file(face["name"], face["bold"],
                                            face["italic"])
        if font_file is None:
            return None, substitutions
        if (face["name"] and used
                and used.casefold() != face["name"].casefold()):
            substitutions.append(f"{face['name']} -> {used}")
        text = "".join(run.get("text", "")
                       for run in paragraph.get("runs", []))
        size = float(face["size"])
        if wrap == "none" or not text.strip():
            lines = 1
        else:
            fitter = HouseTextFitter(
                _LineSource(text),
                (int(avail_w * _EMU_PER_PT), int(avail_h * _EMU_PER_PT)),
                font_file)
            lines = fitter.wrap_line_count(text, max(1, round(size)))
        required += _paragraph_height_pt(paragraph, lines, size)
    return required, substitutions


def assess_frame_record(record: Dict) -> Dict:
    """Fit verdict for one lint-facts shape record (see module doc).

    ``record`` is a ``utils.lint_engine.collect_deck_facts`` shape
    record (resolved paragraphs + the live ``_shape``). Returns at
    least ``{"verdict": ...}``; full assessments add required /
    available heights, the ratio, and any font substitutions.
    """
    shape_elem = record["_shape"]._element
    autofit_kind, _ = autofit_of(shape_elem)
    if autofit_kind == "spAutoFit":
        return {"verdict": "fits", "autofit": "spAutoFit",
                "note": "shape auto-grows to fit its text"}
    box = _frame_box_pt(record)
    if box is None:
        return {"verdict": "unknown",
                "note": "frame has no measurable text box"}
    avail_w, avail_h = box
    required, substitutions = _required_height_pt(
        record, avail_w, avail_h, wrap_of(shape_elem))
    if required is None:
        return {"verdict": "unknown",
                "note": "no font file available to measure with"}
    ratio = required / avail_h
    return {
        "verdict": _verdict(ratio),
        "required_pt": round(required, 1),
        "available_pt": round(avail_h, 1),
        "ratio": round(ratio, 3),
        "autofit": autofit_kind,
        "font_substitutions": sorted(set(substitutions)),
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def predict_text_fit(deck_path: str,
                     slide_number: Optional[int] = None,
                     shape_name: Optional[str] = None) -> Dict:
    """Predict text fit for every text-bearing frame of a deck.

    ``slide_number`` (1-based) and ``shape_name`` narrow the scan.
    Returns ``{"frames": [...], "summary": {verdict: count}}`` where
    each frame entry carries slide / shape refs plus the assessment
    from :func:`assess_frame_record`. Read-only -- the deck is never
    mutated (the whole point vs ``TextFrame.fit_text``).
    """
    from .lint_engine import collect_deck_facts, shape_has_text

    facts = collect_deck_facts(deck_path)
    if slide_number is not None:
        count = len(facts["slides"])
        if not 1 <= int(slide_number) <= count:
            raise ValueError(
                f"slide_number {slide_number!r} outside 1..{count}")
    frames: List[Dict] = []
    summary: Dict[str, int] = {}
    for slide in facts["slides"]:
        if (slide_number is not None
                and slide["slide_number"] != int(slide_number)):
            continue
        for shape in slide["shapes"]:
            if shape_name is not None and shape["name"] != shape_name:
                continue
            if not shape_has_text(shape):
                continue
            result = assess_frame_record(shape)
            frames.append({
                "slide": slide["slide_number"],
                "shape": shape["name"],
                **result,
            })
            verdict = result["verdict"]
            summary[verdict] = summary.get(verdict, 0) + 1
    return {"frames": frames, "summary": summary}
