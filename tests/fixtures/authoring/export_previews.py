"""Export every slide of every fixture deck to PNG contact sheets.

PNGs are written OUTSIDE the repo (default: %TEMP%/ppt_mcp_fixture_previews)
so humans can eyeball the fixtures without bloating version control.
Width is fixed at 1280px; height follows the deck's PageSetup ratio.
"""

from __future__ import annotations

import argparse
import os
import tempfile

from com_helpers import open_presentation, powerpoint_app

FIXTURES = ["theme_only", "layout_override", "explicit_override", "multi_master"]
EXPORT_WIDTH_PX = 1280


def default_output_dir() -> str:
    return os.path.join(tempfile.gettempdir(), "ppt_mcp_fixture_previews")


def export_deck(app, deck_path: str, out_dir: str,
                width_px: int = EXPORT_WIDTH_PX) -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    fixture = os.path.splitext(os.path.basename(deck_path))[0]
    exported: list[str] = []
    with open_presentation(app, deck_path, read_only=True) as pres:
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        height_px = round(width_px * ratio)
        for i in range(1, pres.Slides.Count + 1):
            png_path = os.path.abspath(
                os.path.join(out_dir, f"{fixture}_slide{i:02d}.png"))
            pres.Slides(i).Export(png_path, "PNG", width_px, height_px)
            exported.append(png_path)
    return exported


def export_all(app, fixtures_dir: str, out_dir: str) -> list[str]:
    exported: list[str] = []
    for fixture in FIXTURES:
        deck_path = os.path.join(fixtures_dir, f"{fixture}.pptx")
        paths = export_deck(app, deck_path, out_dir)
        exported.extend(paths)
        print(f"Exported {len(paths)} slides for {fixture}")
    return exported


def main() -> list[str]:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=default_output_dir(),
                        help="output directory for PNGs (kept out of the repo)")
    args = parser.parse_args()
    fixtures_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    with powerpoint_app() as app:
        exported = export_all(app, fixtures_dir, args.out)
    print(f"Wrote {len(exported)} PNGs to {args.out}")
    return exported


if __name__ == "__main__":
    main()
