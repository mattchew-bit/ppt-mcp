"""COM-backed rendering tests (Windows + desktop PowerPoint only).

Marked ``com`` per the repo convention: these skip themselves off-Windows,
when pywin32 is missing, or when desktop PowerPoint is not registered.

Covered here (Verification #4 of the style-fidelity plan):
* render a fixture slide -> PNG exists with dims from the PageSetup ratio
* render succeeds while the SAME deck is open in a WithWindow session
  (temp-copy strategy) -- the test closes only its own presentation
* two consecutive render_deck runs leave no extra POWERPNT processes
* app.AutomationSecurity is restored after a render
"""

import subprocess
import sys
import time

import pytest

from tests.conftest import HOUSE_CORPUS_DIR, fixture_path, skip_if_fixture_missing

FIXTURE = "theme_only.pptx"
HOUSE_DECK = HOUSE_CORPUS_DIR / "house_01.pptx"


def _pywin32_missing():
    try:
        import win32com.client  # noqa: F401
        return False
    except ImportError:
        return True


def _powerpoint_registered():
    if sys.platform != "win32":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"):
            return True
    except OSError:
        return False


pytestmark = [
    pytest.mark.com,
    pytest.mark.skipif(sys.platform != "win32",
                       reason="COM rendering is Windows-only"),
    pytest.mark.skipif(sys.platform == "win32" and _pywin32_missing(),
                       reason="pywin32 not installed"),
    pytest.mark.skipif(sys.platform == "win32" and not _powerpoint_registered(),
                       reason="desktop PowerPoint not registered"),
]


def _powerpnt_count():
    """Number of running POWERPNT.EXE processes."""
    output = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True, text=True,
    ).stdout
    return sum(1 for line in output.splitlines() if "POWERPNT.EXE" in line.upper())


def _wait_for_powerpnt_count(target, timeout_s=20.0):
    deadline = time.monotonic() + timeout_s
    count = _powerpnt_count()
    while count > target and time.monotonic() < deadline:
        time.sleep(0.5)
        count = _powerpnt_count()
    return count


def _expected_height(deck_path, width):
    from pptx import Presentation

    prs = Presentation(str(deck_path))
    return round(width * prs.slide_height / prs.slide_width)


# ------------------------------------------------------------ basic render


@skip_if_fixture_missing(FIXTURE)
def test_render_fixture_slide_expected_dims(tmp_path):
    from PIL import Image

    from utils.render_com import render_slides
    from utils.render_compare import read_renderer_tag

    deck = fixture_path(FIXTURE)
    result = render_slides(str(deck), width=1280, slide_range="1",
                           output_dir=str(tmp_path))

    assert len(result["paths"]) == 1
    png = result["paths"][0]
    expected_h = _expected_height(deck, 1280)
    with Image.open(png) as img:
        assert img.size == (1280, expected_h)
    assert result["width"] == 1280
    assert result["height"] == expected_h
    assert result["renderer"] == "powerpoint"
    assert read_renderer_tag(png) == "powerpoint"


@skip_if_fixture_missing(FIXTURE)
def test_render_custom_width(tmp_path):
    from PIL import Image

    from utils.render_com import render_slides

    deck = fixture_path(FIXTURE)
    result = render_slides(str(deck), width=640, slide_range="1",
                           output_dir=str(tmp_path))

    with Image.open(result["paths"][0]) as img:
        assert img.size == (640, _expected_height(deck, 640))


@skip_if_fixture_missing(FIXTURE)
def test_render_deck_slide_range(tmp_path):
    import os

    from utils.render_com import render_slides

    deck = fixture_path(FIXTURE)
    result = render_slides(str(deck), slide_range="1-2",
                           output_dir=str(tmp_path))

    assert len(result["paths"]) == 2
    assert all(os.path.isfile(p) for p in result["paths"])


@skip_if_fixture_missing(FIXTURE)
def test_render_out_of_range_slide(tmp_path):
    from utils.render_com import render_slides

    deck = fixture_path(FIXTURE)
    with pytest.raises(ValueError):
        render_slides(str(deck), slide_range="999", output_dir=str(tmp_path))


# ------------------------------------- render while the deck is open (user)


@pytest.mark.skipif(not HOUSE_DECK.is_file(),
                    reason="house corpus not present")
def test_render_while_same_deck_open_in_powerpoint(tmp_path):
    """Simulates Matt having the deck open mid-iteration (WithWindow=True).

    The test opens house_01 via its OWN Dispatch first, then render_slides
    must succeed (temp-copy strategy), then the test closes only its own
    presentation -- never anything else in the running instance.
    """
    import pythoncom
    import win32com.client

    from utils.render_com import render_slides

    deck = str(HOUSE_DECK.resolve())

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(deck, False, False, True)  # WithWindow=True
    try:
        result = render_slides(deck, width=960, slide_range="1",
                               output_dir=str(tmp_path))
        assert len(result["paths"]) == 1
        from PIL import Image

        with Image.open(result["paths"][0]) as img:
            assert img.size[0] == 960
    finally:
        pres.Saved = -1  # suppress save prompt
        pres.Close()
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
        del pres, app
        pythoncom.CoUninitialize()


# ----------------------------------------------------------- process hygiene


@skip_if_fixture_missing(FIXTURE)
def test_consecutive_render_deck_runs_no_zombie_processes(tmp_path):
    from utils.render_com import render_slides

    deck = str(fixture_path(FIXTURE))
    baseline = _powerpnt_count()

    render_slides(deck, slide_range="1-2", output_dir=str(tmp_path / "run1"))
    render_slides(deck, slide_range="1-2", output_dir=str(tmp_path / "run2"))

    assert _wait_for_powerpnt_count(baseline) <= baseline


@skip_if_fixture_missing(FIXTURE)
def test_automation_security_restored_after_render(tmp_path):
    import pythoncom
    import win32com.client

    from utils.render_com import render_slides

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        before = app.AutomationSecurity

        render_slides(str(fixture_path(FIXTURE)), slide_range="1",
                      output_dir=str(tmp_path))

        assert app.AutomationSecurity == before
    finally:
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
        del app
        pythoncom.CoUninitialize()


# --------------------------------------------- launch-ownership kill switch


@skip_if_fixture_missing(FIXTURE)
def test_launched_powerpnt_pid_is_recorded(tmp_path, monkeypatch):
    """Regression: ``app.HWND`` raises DISP_E_MEMBERNOTFOUND on late-bound
    Dispatch (Office16), so record_launch always received None and the
    emergency kill switch was permanently disarmed. The PID must now come
    from the process-list diff around Dispatch."""
    from utils import render_com

    if _wait_for_powerpnt_count(0) > 0:
        pytest.skip("PowerPoint already running; launch path not exercised")

    captured = []
    original = render_com._WORKER.record_launch

    def spy(pid):
        captured.append(pid)
        original(pid)

    monkeypatch.setattr(render_com._WORKER, "record_launch", spy)

    render_com.render_slides(str(fixture_path(FIXTURE)), slide_range="1",
                             output_dir=str(tmp_path))

    assert captured, "launch path never recorded a PID"
    assert captured[0] is not None, (
        "launched-PID discovery returned None -- the emergency kill "
        "switch would be disarmed for this launch")
    assert isinstance(captured[0], int) and captured[0] > 0


# --------------------------------------------- corrupt deck at the COM layer


def test_corrupt_zip_maps_to_actionable_error(tmp_path):
    """Regression: a corrupt file with a valid PK signature reached
    Presentations.Open and surfaced as a raw com_error hex tuple. It must
    now come back as an actionable ValueError."""
    from utils.render_com import render_slides

    crafted = tmp_path / "corrupt.pptx"
    crafted.write_bytes(b"PK\x03\x04" + b"\x99" * 4096)

    with pytest.raises(ValueError) as excinfo:
        render_slides(str(crafted), slide_range="1",
                      output_dir=str(tmp_path / "out"))

    message = str(excinfo.value)
    assert "corrupt.pptx" in message
    assert "corrupt" in message.lower()
    assert "-2147352567" not in message.split("(COM detail")[0]


# --------------------------------------------------------- tool end-to-end


@skip_if_fixture_missing(FIXTURE)
def test_render_slide_tool_end_to_end():
    from mcp.server.fastmcp import Image as FastMCPImage

    from tools.render_tools import register_render_tools

    class _RecorderApp:
        def __init__(self):
            self.tools = {}

        def tool(self, *args, **kwargs):
            def decorator(fn):
                self.tools[fn.__name__] = fn
                return fn

            return decorator

    app = _RecorderApp()
    register_render_tools(app)

    result = app.tools["render_slide"](
        file_path=str(fixture_path(FIXTURE)), slide_index=0)

    assert isinstance(result, list), f"expected envelope list, got {result!r}"
    envelope, image = result[0], result[1]
    assert envelope["renderer"] == "powerpoint"
    assert envelope["width"] == 1280
    assert isinstance(image, FastMCPImage)
