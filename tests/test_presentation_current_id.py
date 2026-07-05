"""Regression tests: create/open set the CURRENT presentation id.

``open_presentation`` (and the create tools) stored the deck but never
set the current presentation id, so any tool whose ``presentation_id``
defaults to "the current presentation" (documented ``(default:
current)``) returned "No presentation loaded" in the exact
create/open -> apply flow the docs promote.

``register_presentation_tools`` now takes an optional
``set_current_presentation_id`` callable and invokes it whenever a deck
is stored; the server wires its module-level setter through. These
tests pin the setter contract and the open -> apply-by-default flow.
"""

import copy
import json
import shutil

import pytest
from pptx import Presentation

from tests.conftest import fixture_path
from tests.test_apply_house_profile import (
    DEVIANT,
    HOUSE_PROFILE,
    _RecorderApp,
)
from tools.presentation_tools import register_presentation_tools
from tools.style_tools import register_style_tools


class _CurrentId:
    """Mutable current-presentation-id holder mimicking the server's."""

    def __init__(self):
        self.value = None

    def get(self):
        return self.value

    def set(self, pres_id):
        self.value = pres_id


def _presentation_tools(presentations, current, with_setter=True):
    app = _RecorderApp()
    if with_setter:
        register_presentation_tools(app, presentations, current.get,
                                    lambda: [], current.set)
    else:
        register_presentation_tools(app, presentations, current.get,
                                    lambda: [])
    return app.tools


@pytest.fixture
def deviant_copy(tmp_path):
    source = fixture_path(DEVIANT)
    if not source.is_file():
        pytest.skip(f"fixture {DEVIANT} not present")
    copy_path = tmp_path / "deviant.pptx"
    shutil.copyfile(source, copy_path)
    return copy_path


def test_open_presentation_sets_current_id(deviant_copy):
    presentations, current = {}, _CurrentId()
    tools = _presentation_tools(presentations, current)
    result = tools["open_presentation"](file_path=str(deviant_copy))
    assert "error" not in result
    assert result["presentation_id"] in presentations
    assert current.get() == result["presentation_id"]


def test_create_presentation_sets_current_id():
    presentations, current = {}, _CurrentId()
    tools = _presentation_tools(presentations, current)
    result = tools["create_presentation"]()
    assert current.get() == result["presentation_id"]


def test_register_without_setter_stays_backward_compatible(deviant_copy):
    presentations, current = {}, _CurrentId()
    tools = _presentation_tools(presentations, current, with_setter=False)
    result = tools["open_presentation"](file_path=str(deviant_copy))
    assert "error" not in result
    assert result["presentation_id"] in presentations
    assert current.get() is None  # legacy behavior: caller manages it


def test_open_then_apply_house_profile_without_explicit_id(
        tmp_path, deviant_copy):
    """The documented flow: open a deck, then apply the loaded house
    profile WITHOUT passing presentation_id."""
    presentations, current = {}, _CurrentId()
    tools = _presentation_tools(presentations, current)

    style_app = _RecorderApp()
    register_style_tools(style_app, presentations, current.get)

    profile = copy.deepcopy(HOUSE_PROFILE)
    profile["name"] = "meridian_current_id_flow"
    profile_path = tmp_path / "profile.json"
    profile_path.write_text(json.dumps(profile), encoding="utf-8")
    loaded = style_app.tools["load_style_profile"](
        file_path=str(profile_path))
    assert "error" not in loaded

    opened = tools["open_presentation"](file_path=str(deviant_copy))
    assert "error" not in opened

    result = style_app.tools["apply_style_profile"](
        profile_name="meridian_current_id_flow")
    assert "error" not in result
    assert result["writes"] == 7

    saved = tools["save_presentation"](
        file_path=str(tmp_path / "saved.pptx"))
    assert "error" not in saved
    assert Presentation(str(tmp_path / "saved.pptx")).slides is not None
