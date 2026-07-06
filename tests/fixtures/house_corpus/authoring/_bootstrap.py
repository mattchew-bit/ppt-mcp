"""Path bootstrap for the house-corpus authoring scripts.

Importing this module makes two other source roots importable:

* the repo root                  -> ``utils.*`` (resolver modules)
* ``tests/fixtures/authoring``   -> ``com_helpers`` / ``extract_expected``
  (the Step 0 fixture-authoring toolkit, reused verbatim -- same COM
  hygiene: CoInitialize, never Visible=False, own presentations only,
  Quit only when Presentations.Count == 0, absolute paths, SaveAs 24)

It also exposes the corpus directory layout as absolute paths.
"""

from __future__ import annotations

import os
import sys

AUTHORING_DIR = os.path.dirname(os.path.abspath(__file__))
CORPUS_DIR = os.path.dirname(AUTHORING_DIR)
FIXTURES_DIR = os.path.dirname(CORPUS_DIR)
TESTS_DIR = os.path.dirname(FIXTURES_DIR)
REPO_ROOT = os.path.dirname(TESTS_DIR)

FIXTURE_AUTHORING_DIR = os.path.join(FIXTURES_DIR, "authoring")
ASSETS_DIR = os.path.join(AUTHORING_DIR, "assets")
EXPECTED_DIR = os.path.join(CORPUS_DIR, "expected_values")

for _path in (REPO_ROOT, FIXTURE_AUTHORING_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)


def corpus_path(filename: str) -> str:
    """Absolute path of a file inside tests/fixtures/house_corpus/."""
    return os.path.join(CORPUS_DIR, filename)


def asset_path(filename: str) -> str:
    """Absolute path of a generated image asset."""
    return os.path.join(ASSETS_DIR, filename)
