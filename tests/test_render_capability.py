"""Capability / pre-flight tests for the rendering stack (non-COM).

Covers the guardrails that must work on ANY machine, including ones without
pywin32 or LibreOffice:

* encrypted-deck magic-byte pre-check (OLE header => refuse before COM)
* lazy-import capability errors when pywin32 is missing (simulated by
  poisoning ``sys.modules`` -- a ``None`` entry makes ``import`` raise)
* LibreOffice capability error when ``soffice`` is absent
* renderer auto-selection falls back COM -> LibreOffice -> combined error
"""

import sys

import pytest

from utils.render_com import (
    RenderCapabilityError,
    check_not_encrypted,
    ensure_com_capability,
    validate_render_source,
)


OLE_MAGIC = b"\xd0\xcf\x11\xe0"


# ------------------------------------------------- encrypted-deck pre-check


def _write_bytes(path, data):
    with open(str(path), "wb") as handle:
        handle.write(data)
    return str(path)


def test_ole_magic_bytes_refused(tmp_path):
    crafted = _write_bytes(tmp_path / "locked.pptx",
                           OLE_MAGIC + b"\xa1\xb1\x1a\xe1" + b"\x00" * 64)

    with pytest.raises(ValueError) as excinfo:
        check_not_encrypted(crafted)

    message = str(excinfo.value)
    assert "locked.pptx" in message
    assert "password" in message.lower()


def test_zip_magic_bytes_accepted(tmp_path):
    crafted = _write_bytes(tmp_path / "plain.pptx", b"PK\x03\x04" + b"\x00" * 64)
    check_not_encrypted(crafted)  # must not raise


def test_truncated_zip_refused(tmp_path):
    """Regression: sub-minimum files used to sail through to COM."""
    crafted = _write_bytes(tmp_path / "tiny.pptx", b"PK")

    with pytest.raises(ValueError) as excinfo:
        check_not_encrypted(crafted)

    assert "corrupt" in str(excinfo.value).lower()


def test_zero_byte_file_refused(tmp_path):
    """Regression: PowerPoint opens a zero-byte .pptx as a blank 0-slide
    presentation, so render_deck returned a silent success envelope."""
    crafted = _write_bytes(tmp_path / "empty.pptx", b"")

    with pytest.raises(ValueError) as excinfo:
        check_not_encrypted(crafted)

    message = str(excinfo.value)
    assert "empty.pptx" in message
    assert "empty" in message.lower()


def test_non_zip_garbage_refused(tmp_path):
    """Regression: garbage bytes used to reach COM and surface as a raw
    com_error hex tuple at the tool boundary."""
    crafted = _write_bytes(tmp_path / "garbage.pptx", b"not a zip at all" * 8)

    with pytest.raises(ValueError) as excinfo:
        check_not_encrypted(crafted)

    message = str(excinfo.value)
    assert "garbage.pptx" in message
    assert "corrupt" in message.lower()


def test_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        check_not_encrypted(str(tmp_path / "absent.pptx"))


def test_validate_render_source_rejects_encrypted(tmp_path):
    crafted = _write_bytes(tmp_path / "locked.pptx", OLE_MAGIC + b"\x00" * 64)
    with pytest.raises(ValueError):
        validate_render_source(crafted)


def test_validate_render_source_rejects_extension(tmp_path):
    crafted = _write_bytes(tmp_path / "notes.txt", b"hello")
    with pytest.raises(ValueError) as excinfo:
        validate_render_source(crafted)
    assert ".txt" in str(excinfo.value)


# ------------------------------------------------ pywin32 capability errors


def _poison_pywin32(monkeypatch):
    """Simulate a machine without pywin32: None entries make import raise."""
    for name in ("win32com", "win32com.client", "pythoncom", "win32process"):
        monkeypatch.setitem(sys.modules, name, None)


def test_com_capability_error_without_pywin32(monkeypatch):
    _poison_pywin32(monkeypatch)

    with pytest.raises(RenderCapabilityError) as excinfo:
        ensure_com_capability()

    message = str(excinfo.value)
    assert "ppt-mcp[render]" in message
    assert "PowerPoint" in message


def test_com_capability_error_without_powerpoint_registration(monkeypatch):
    """Regression: pywin32 present but PowerPoint absent used to pass the
    capability check and then leak a temp deck copy per render call."""
    if sys.platform != "win32":
        pytest.skip("registry-based PowerPoint check is Windows-only")
    pytest.importorskip("win32com.client", reason="pywin32 not installed")
    import winreg

    def refuse_key(*args, **kwargs):
        raise OSError("simulated: ProgID not registered")

    monkeypatch.setattr(winreg, "OpenKey", refuse_key)

    with pytest.raises(RenderCapabilityError) as excinfo:
        ensure_com_capability()

    assert "not registered" in str(excinfo.value)


def test_com_capability_error_off_windows(monkeypatch):
    monkeypatch.setattr(sys, "platform", "linux")

    with pytest.raises(RenderCapabilityError) as excinfo:
        ensure_com_capability()

    assert "Windows" in str(excinfo.value)


def test_module_imports_without_pywin32():
    """utils.render_com must be importable with pywin32 absent (lazy imports).

    Runs in a subprocess so the in-process module graph is never disturbed
    (an in-process reload would split class identity for
    RenderCapabilityError across modules).
    """
    import os
    import subprocess

    code = (
        "import sys\n"
        "for name in ('win32com', 'win32com.client', 'pythoncom',"
        " 'win32process'):\n"
        "    sys.modules[name] = None\n"
        "import utils.render_com\n"
        "import utils.render_lo\n"
        "import tools.render_tools\n"
        "import ppt_mcp_server\n"
        "print('IMPORT_OK')\n"
    )
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    completed = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True, text=True, cwd=repo_root, timeout=120,
    )
    assert completed.returncode == 0, completed.stderr
    assert "IMPORT_OK" in completed.stdout


# ------------------------------------------------- LibreOffice capability


def test_lo_capability_error_without_soffice(monkeypatch):
    from utils import render_lo

    monkeypatch.setattr(render_lo, "find_soffice", lambda: None)

    with pytest.raises(RenderCapabilityError) as excinfo:
        render_lo.ensure_lo_capability()

    assert "LibreOffice" in str(excinfo.value)


# ------------------------------------------------- renderer auto-selection


def test_select_renderer_combined_error(monkeypatch):
    from tools import render_tools
    from utils import render_lo

    _poison_pywin32(monkeypatch)
    monkeypatch.setattr(render_lo, "find_soffice", lambda: None)

    with pytest.raises(RenderCapabilityError) as excinfo:
        render_tools._select_renderer()

    message = str(excinfo.value)
    assert "ppt-mcp[render]" in message
    assert "LibreOffice" in message


def test_select_renderer_falls_back_to_libreoffice(monkeypatch):
    from tools import render_tools
    from utils import render_lo

    _poison_pywin32(monkeypatch)
    monkeypatch.setattr(render_lo, "find_soffice", lambda: "C:/fake/soffice.exe")

    assert render_tools._select_renderer() == "libreoffice"


def test_select_renderer_prefers_powerpoint():
    pytest.importorskip("win32com.client",
                        reason="pywin32 not installed on this machine")
    if sys.platform != "win32":
        pytest.skip("COM renderer is Windows-only")
    from tools import render_tools

    assert render_tools._select_renderer() == "powerpoint"
