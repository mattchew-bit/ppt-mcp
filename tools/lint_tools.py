"""Style-lint and text-fit tools for PowerPoint MCP Server (Step 5).

MCP-facing wrappers over ``utils.lint_engine`` (deck-vs-house-profile
conformance), ``utils.text_fit`` (standalone overflow prediction) and
the ``diff_decks`` reuse of the engine with deck-B-as-profile
semantics. Responses are hard-capped (~40KB, like the resolved
analyzer) with explicit truncation markers.
"""

import json
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

#: Hard cap for serialized lint/fit responses (analyzer convention).
MAX_LINT_RESPONSE_BYTES = 40_000

_VALID_FLOORS = ("error", "warn", "info")


def _apply_severity_floor(findings: List[Dict],
                          severity_floor: Optional[str]) -> List[Dict]:
    if severity_floor is None:
        return findings
    if severity_floor not in _VALID_FLOORS:
        raise ValueError(
            f"severity_floor must be one of {_VALID_FLOORS}, got "
            f"{severity_floor!r}"
        )
    from utils.lint_engine import SEVERITY_RANK

    floor_rank = SEVERITY_RANK[severity_floor]
    return [f for f in findings
            if SEVERITY_RANK[f["severity"]] <= floor_rank]


def _apply_slide_range(findings: List[Dict], slide_range: Optional[str],
                       slide_count: int) -> List[Dict]:
    if not slide_range:
        return findings
    from utils.resolve_analysis import parse_slide_range

    selected = {index + 1
                for index in parse_slide_range(slide_range, slide_count)}
    return [f for f in findings if f["slide"] in selected]


def _summarize(findings: List[Dict]) -> Dict:
    by_severity: Dict[str, int] = {}
    by_rule: Dict[str, int] = {}
    for finding in findings:
        by_severity[finding["severity"]] = (
            by_severity.get(finding["severity"], 0) + 1)
        by_rule[finding["rule_id"]] = by_rule.get(finding["rule_id"],
                                                  0) + 1
    return {"total": len(findings), "by_severity": by_severity,
            "by_rule": by_rule}


def _cap_response(result: Dict, key: str,
                  max_bytes: Optional[int] = None) -> Dict:
    """Drop trailing ``result[key]`` entries until the response fits."""
    if max_bytes is None:  # resolved at call time (testable)
        max_bytes = MAX_LINT_RESPONSE_BYTES
    entries = list(result[key])
    total = len(entries)
    while entries and len(json.dumps(result,
                                     ensure_ascii=False)) > max_bytes:
        entries.pop()
        result[key] = entries
        result["truncated"] = True
        result["hint"] = (
            f"response capped at {max_bytes} bytes: {len(entries)} of "
            f"{total} {key} returned -- narrow with slide_range or "
            "severity_floor"
        )
    return result


def _slide_count(file_path: str) -> int:
    from pptx import Presentation

    return len(Presentation(file_path).slides)


def register_lint_tools(app: FastMCP):
    """Register style-lint and text-fit tools with the FastMCP app."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Lint Deck Against House Profile",
            readOnlyHint=True,
        ),
    )
    def lint_against_profile(
        file_path: str,
        profile_name: str,
        severity_floor: Optional[str] = None,
        slide_range: Optional[str] = None,
    ) -> Dict:
        """Check a deck against a learned house-style profile.

        Compares the inheritance-RESOLVED deck analysis (effective
        values, not just explicit XML) with a house-profile/1 profile
        created by create_house_profile (or loaded via
        load_style_profile) and returns an ordered deviation list.
        Rule catalog v1: font-scale, font-family (latin/ea/cs/sym +
        buFont coverage), bullet-style, spacing, color-palette,
        hardcoded-color, border-style, off-grid (distance to the
        learned alignment grid), straggler-textbox, off-slide,
        archetype-geometry, autofit-shrink, text-overflow-predicted,
        proofing-language, image-distortion, image-dpi,
        footer-presence, overlap, empty-slide, tiny-font.

        Findings are ordered by severity (error/warn/info) then slide;
        each carries slide/shape (and paragraph/run where applicable),
        property, expected, actual, and a distribution-style message
        (e.g. "body run is 13pt; house type scale is {11, 14, 20,
        30}pt"). Deterministic: rerun after fixes until findings reach
        zero (the generate -> lint -> fix loop). Read-only; the
        response is capped (~40KB) with a truncation marker.

        Args:
            file_path: Path to the .pptx deck to lint
            profile_name: Name of a loaded house-profile/1 profile
                (see create_house_profile / load_style_profile)
            severity_floor: Only return findings at or above this
                severity: "error", "warn" or "info" (default: all)
            slide_range: 1-based slides to report, e.g. "1-3,5"
                (default: all slides)
        """
        from pathlib import Path

        from tools import style_tools
        from utils.lint_engine import lint_against_profile as run_lint
        from utils.style_apply import is_house_profile

        profile = style_tools._style_profiles.get(profile_name)
        if profile is None:
            available = (list(style_tools._style_profiles.keys())
                         or ["(none)"])
            return {"error": f"Profile '{profile_name}' not found. "
                             f"Available: {', '.join(available)}"}
        if not is_house_profile(profile):
            return {"error": f"Profile '{profile_name}' is not a "
                             "house-profile/1 profile -- lint needs one "
                             "created by create_house_profile"}
        try:
            if not Path(file_path).is_file():
                raise FileNotFoundError(f"deck not found: {file_path}")
            findings = run_lint(file_path, profile)
            findings = _apply_severity_floor(findings, severity_floor)
            findings = _apply_slide_range(findings, slide_range,
                                          _slide_count(file_path))
        except FileNotFoundError as e:
            return {"error": str(e)}
        except ValueError as e:
            return {"error": f"Invalid lint options: {str(e)}"}
        except Exception as e:
            return {"error": f"Lint failed: {str(e)}"}

        result = {
            "message": (
                f"Lint of {file_path} against '{profile_name}': "
                f"{len(findings)} finding(s)"
            ),
            "profile_name": profile_name,
            "summary": _summarize(findings),
            "truncated": False,
            "findings": findings,
        }
        return _cap_response(result, "findings")

    @app.tool(
        annotations=ToolAnnotations(
            title="Predict Text Fit",
            readOnlyHint=True,
        ),
    )
    def predict_text_fit(
        file_path: str,
        slide_index: Optional[int] = None,
        shape_name: Optional[str] = None,
    ) -> Dict:
        """Predict whether each text frame's declared text fits its box.

        Runs python-pptx's TextFitter machinery STANDALONE (the deck is
        never mutated) with per-paragraph inheritance-resolved sizes, a
        realistic ~1.2x line pitch honoring lnSpc overrides, and a font
        resolver covering the system and per-user font directories plus
        the registry (Arial/Calibri fallback, substitutions reported).

        Verdict per frame: "fits", "overflow", or "borderline" (within
        ±5% of the frame -- confirm by render); frames that auto-grow
        (spAutoFit) report "fits", while normAutofit frames are still
        predicted -- python-pptx-written decks carry stale normAutofit
        scale factors, and renderers draw that overflow as-is.

        Args:
            file_path: Path to the .pptx deck to check
            slide_index: Index of the slide to check (0-based,
                default: all slides)
            shape_name: Only check the frame with this shape name
        """
        from pathlib import Path

        from utils.text_fit import predict_text_fit as run_predict

        try:
            if not Path(file_path).is_file():
                raise FileNotFoundError(f"deck not found: {file_path}")
            slide_number = (None if slide_index is None
                            else int(slide_index) + 1)
            report = run_predict(file_path, slide_number=slide_number,
                                 shape_name=shape_name)
        except FileNotFoundError as e:
            return {"error": str(e)}
        except ValueError as e:
            return {"error": f"Invalid text-fit options: {str(e)}"}
        except Exception as e:
            return {"error": f"Text-fit prediction failed: {str(e)}"}

        result = {
            "message": (
                f"Text-fit prediction for {file_path}: "
                f"{report['summary']} across {len(report['frames'])} "
                "frame(s)"
            ),
            "summary": report["summary"],
            "truncated": False,
            "frames": report["frames"],
        }
        return _cap_response(result, "frames")

    @app.tool(
        annotations=ToolAnnotations(
            title="Diff Two Decks",
            readOnlyHint=True,
        ),
    )
    def diff_decks(
        file_path_a: str,
        file_path_b: str,
        severity_floor: Optional[str] = None,
    ) -> Dict:
        """Report where deck A deviates from deck B's style conventions.

        Reuses the lint engine with deck-B-as-profile semantics: a
        house-profile/1 profile is learned from deck B alone, then deck
        A is linted against it. Findings read "A deviates from B's
        convention" -- asymmetric by design (swap the arguments for the
        other direction). Both decks must share one slide size.

        Caveat: conventions learned from a SINGLE deck are weaker than
        a real house profile (grid lines need multi-slide support), so
        geometry findings are candidates to review, not verdicts; the
        deterministic style rules (fonts, sizes, colors, spacing,
        bullets, borders) compare exactly.

        Args:
            file_path_a: Path of the deck to check (the "candidate")
            file_path_b: Path of the deck whose style is the reference
            severity_floor: Only return findings at or above this
                severity: "error", "warn" or "info" (default: all)
        """
        from pathlib import Path

        from utils.lint_engine import lint_against_profile as run_lint
        from utils.profile_extract import create_house_profile as build

        try:
            for path in (file_path_a, file_path_b):
                if not Path(path).is_file():
                    raise FileNotFoundError(f"deck not found: {path}")
            profile = build([file_path_b], "deck_b_reference")
            findings = run_lint(file_path_a, profile)
            findings = _apply_severity_floor(findings, severity_floor)
        except FileNotFoundError as e:
            return {"error": str(e)}
        except ValueError as e:
            return {"error": f"diff_decks could not build a reference "
                             f"profile from {file_path_b}: {str(e)}"}
        except Exception as e:
            return {"error": f"Deck diff failed: {str(e)}"}

        result = {
            "message": (
                f"{Path(file_path_a).name} vs "
                f"{Path(file_path_b).name} (as reference): "
                f"{len(findings)} deviation(s)"
            ),
            "reference_deck": file_path_b,
            "summary": _summarize(findings),
            "truncated": False,
            "findings": findings,
        }
        return _cap_response(result, "findings")
