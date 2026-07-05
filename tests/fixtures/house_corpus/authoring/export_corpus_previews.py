"""Export contact-sheet PNGs of every house-corpus deck (COM).

PNGs are written OUTSIDE the repo (they are eyeball artifacts, not
fixtures). Width fixed at 1280px, height follows the deck ratio.

Named export_corpus_previews (not export_previews) on purpose: the
Step 0 fixture kit on sys.path already owns ``export_previews`` and
would shadow a same-named module here.
"""

from __future__ import annotations

import argparse
import os
import tempfile

import _bootstrap
from com_helpers import open_presentation, powerpoint_app

DECKS = ("house_01", "house_02", "house_03", "house_04", "house_05",
         "deviant_01")
EXPORT_WIDTH_PX = 1280


def default_output_dir() -> str:
    return os.path.join(tempfile.gettempdir(), "house_corpus_previews")


def export_deck(app, deck_name: str, out_dir: str) -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    deck_path = _bootstrap.corpus_path(f"{deck_name}.pptx")
    exported: list[str] = []
    with open_presentation(app, deck_path, read_only=True) as pres:
        ratio = pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth
        height_px = round(EXPORT_WIDTH_PX * ratio)
        for i in range(1, pres.Slides.Count + 1):
            png_path = os.path.abspath(
                os.path.join(out_dir, f"{deck_name}_slide{i:02d}.png"))
            pres.Slides(i).Export(png_path, "PNG", EXPORT_WIDTH_PX,
                                  height_px)
            exported.append(png_path)
    return exported


def export_all(out_dir: str) -> list[str]:
    exported: list[str] = []
    with powerpoint_app() as app:
        for deck_name in DECKS:
            paths = export_deck(app, deck_name, out_dir)
            exported.extend(paths)
            print(f"Exported {len(paths)} slides for {deck_name}")
    print(f"Wrote {len(exported)} PNGs to {out_dir}")
    return exported


def main() -> list[str]:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", default=default_output_dir(),
                        help="output directory for PNGs (outside the repo)")
    args = parser.parse_args()
    return export_all(args.out)


if __name__ == "__main__":
    main()
