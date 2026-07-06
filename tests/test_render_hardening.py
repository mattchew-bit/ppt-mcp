"""Regression tests for the render-stack repair pass (non-COM).

Each test pins one audited defect in ``utils/render_com.py``:

* temp deck copies (``ppt_mcp_render_src_*``) must never leak, even when
  attach/launch or the copy itself fails (confidential client content)
* ``_emergency_cleanup`` must verify the recorded PID still belongs to a
  POWERPNT.EXE and must report a kill ONLY when taskkill succeeded
* a job's timeout budget must start when the job is dequeued, so a quick
  render queued behind a slow-but-healthy one neither times out spuriously
  nor taskkills the healthy render's PowerPoint
* stale render temp dirs from crashed/killed sessions are age-swept

All tests here run on any platform: the worker loop is exercised without
``pythoncom`` via a subclass that skips COM initialization.
"""

import os
import subprocess
import tempfile
import threading
import time
from types import SimpleNamespace

import pytest

from utils import render_com
from utils.render_com import RenderCapabilityError, RenderTimeoutError


# ------------------------------------------------- temp deck-copy leaks


def _src_dir_names():
    try:
        return {name for name in os.listdir(tempfile.gettempdir())
                if name.startswith(render_com.TEMP_SRC_PREFIX)}
    except OSError:
        return set()


def _fake_deck(tmp_path, name="deck.pptx"):
    deck = tmp_path / name
    deck.write_bytes(b"PK\x03\x04" + b"\x00" * 64)
    return str(deck)


def test_render_job_cleans_deck_copy_when_launch_fails(tmp_path, monkeypatch):
    """Regression: mkdtemp+copy2 ran before the try/finally, so a failing
    attach/launch (e.g. pywin32 without PowerPoint, or a transient
    CO_E_SERVER_EXEC_FAILURE) leaked a full deck copy in %TEMP%."""
    deck = _fake_deck(tmp_path)

    def explode():
        raise RenderCapabilityError("simulated: PowerPoint unavailable")

    monkeypatch.setattr(render_com, "_attach_or_launch", explode)
    before = _src_dir_names()

    with pytest.raises(RenderCapabilityError):
        render_com._render_job(deck, None, 1280, str(tmp_path / "out"))

    assert _src_dir_names() == before


def test_render_job_cleans_temp_dir_when_copy_fails(tmp_path):
    """Same leak family: the deck vanishing between validation and the
    worker picking the job up must not strand the fresh temp dir."""
    before = _src_dir_names()

    with pytest.raises(FileNotFoundError):
        render_com._render_job(str(tmp_path / "vanished.pptx"),
                               None, 1280, str(tmp_path / "out"))

    assert _src_dir_names() == before


# ---------------------------------------------------- emergency cleanup


def _worker_with_pid(pid):
    worker = render_com._ComWorker()
    worker.record_launch(pid)
    return worker


def _stub_subprocess(monkeypatch, calls, returncode=0):
    """Replace render_com's subprocess with a recording stub."""

    def fake_run(cmd, **kwargs):
        calls.append(list(cmd))
        return SimpleNamespace(returncode=returncode, stdout="", stderr="")

    stub = SimpleNamespace(run=fake_run,
                           SubprocessError=subprocess.SubprocessError)
    monkeypatch.setattr(render_com, "subprocess", stub)


def test_emergency_cleanup_stale_pid_not_killed_and_forgotten(monkeypatch):
    """Regression: a stale/reused PID was taskkilled unconditionally and
    the cleanup still claimed success ('was terminated')."""
    calls = []
    _stub_subprocess(monkeypatch, calls)
    monkeypatch.setattr(render_com, "_powerpnt_pids", lambda: set())
    worker = _worker_with_pid(os.getpid())  # a real, non-POWERPNT pid

    assert worker._emergency_cleanup() is False
    assert calls == []  # no taskkill against a non-POWERPNT process
    assert worker._launched_pid is None  # stale record forgotten


def test_emergency_cleanup_failed_taskkill_reports_false_keeps_pid(
        monkeypatch):
    """Regression: a failed taskkill still returned True AND cleared the
    pid, permanently disarming retries against the genuinely hung
    process."""
    calls = []
    _stub_subprocess(monkeypatch, calls, returncode=1)
    monkeypatch.setattr(render_com, "_powerpnt_pids", lambda: {4242})
    worker = _worker_with_pid(4242)

    assert worker._emergency_cleanup() is False
    assert len(calls) == 1
    assert worker._launched_pid == 4242  # kept for the next timeout


def test_emergency_cleanup_success_kills_and_clears(monkeypatch):
    calls = []
    _stub_subprocess(monkeypatch, calls, returncode=0)
    monkeypatch.setattr(render_com, "_powerpnt_pids", lambda: {4242})
    worker = _worker_with_pid(4242)

    assert worker._emergency_cleanup() is True
    assert calls == [["taskkill", "/PID", "4242", "/T", "/F"]]
    assert worker._launched_pid is None


def test_emergency_cleanup_foreign_presentations_block_kill(monkeypatch):
    calls = []
    _stub_subprocess(monkeypatch, calls)
    monkeypatch.setattr(render_com, "_powerpnt_pids", lambda: {4242})
    worker = _worker_with_pid(4242)
    worker.record_foreign_presentations(True)

    assert worker._emergency_cleanup() is False
    assert calls == []


# --------------------------------------------------- timeout semantics


class _PlainWorker(render_com._ComWorker):
    """Worker whose thread runs the job loop without COM initialization."""

    def _main(self):
        self._run_jobs()


def _submit_in_background(worker, fn, timeout):
    thread = threading.Thread(
        target=lambda: worker.submit(fn, timeout=timeout), daemon=True)
    thread.start()
    return thread


def test_timeout_budget_starts_when_job_is_dequeued():
    """Regression: the budget used to start at submit, so a quick job
    behind a slow-but-healthy one reliably timed out through no fault
    of its own."""
    worker = _PlainWorker()
    slow = _submit_in_background(worker, lambda: time.sleep(1.2), timeout=10)
    time.sleep(0.1)  # ensure the slow job is dequeued first

    # Queue wait (~1.1s) + own run (1.2s) = ~2.3s total > 2.0s timeout:
    # the old submit-anchored budget failed this; per-phase budget passes.
    result = worker.submit(lambda: time.sleep(1.2) or "ok", timeout=2.0)

    assert result == "ok"
    slow.join(timeout=5)


def test_queue_timeout_never_triggers_emergency_cleanup(monkeypatch):
    """Regression: a job that timed out while still QUEUED used to run
    emergency cleanup, taskkilling the POWERPNT mid-export for the
    healthy job in front of it."""
    worker = _PlainWorker()
    cleanup_calls = []
    monkeypatch.setattr(worker, "_emergency_cleanup",
                        lambda: cleanup_calls.append(True) or True)
    slow = _submit_in_background(worker, lambda: time.sleep(1.0), timeout=10)
    time.sleep(0.1)

    with pytest.raises(RenderTimeoutError) as excinfo:
        worker.submit(lambda: "never runs", timeout=0.3)

    message = str(excinfo.value)
    assert "waiting for the render worker" in message
    assert "Nothing was terminated" in message
    assert cleanup_calls == []
    slow.join(timeout=5)


def test_queue_timed_out_job_is_cancelled_not_executed():
    """A caller that gave up waiting must not have its job run later
    (it would render into an output dir nobody will ever read)."""
    worker = _PlainWorker()
    ran = []
    slow = _submit_in_background(worker, lambda: time.sleep(0.8), timeout=10)
    time.sleep(0.1)

    with pytest.raises(RenderTimeoutError):
        worker.submit(lambda: ran.append(True), timeout=0.2)

    slow.join(timeout=5)
    time.sleep(0.3)  # give the worker time to drain the cancelled job
    assert ran == []


def test_run_timeout_message_is_truthful_when_nothing_was_killed(
        monkeypatch):
    """Regression: the timeout text claimed 'this server did not launch
    it' even when we did launch it and the kill switch simply failed."""
    worker = _PlainWorker()
    monkeypatch.setattr(worker, "_emergency_cleanup", lambda: False)

    with pytest.raises(RenderTimeoutError) as excinfo:
        worker.submit(lambda: time.sleep(2.0), timeout=0.3)

    message = str(excinfo.value)
    assert "was terminated" not in message
    assert "PowerPoint was left running" in message
    assert "POWERPNT.EXE" in message


# -------------------------------------------------------- stale sweep


def _make_dir(root, name, age_s):
    path = root / name
    path.mkdir()
    (path / "payload.bin").write_bytes(b"x" * 16)
    stamp = time.time() - age_s
    os.utime(str(path), (stamp, stamp))
    return str(path)


def test_sweep_reaps_only_stale_render_dirs(tmp_path, monkeypatch):
    """Regression: orphaned deck copies and default output dirs
    accumulated in %TEMP% forever (Windows never purges it)."""
    monkeypatch.setattr(render_com.tempfile, "gettempdir",
                        lambda: str(tmp_path))
    hour = 60.0 * 60.0
    stale_src = _make_dir(tmp_path, "ppt_mcp_render_src_dead", 2 * hour)
    fresh_src = _make_dir(tmp_path, "ppt_mcp_render_src_live", 0)
    stale_lo = _make_dir(tmp_path, "ppt_mcp_lo_dead", 2 * hour)
    stale_out = _make_dir(tmp_path, "ppt_mcp_render_deck_old", 192 * hour)
    fresh_out = _make_dir(tmp_path, "ppt_mcp_render_deck_new", 24 * hour)
    unrelated = _make_dir(tmp_path, "unrelated_dir", 500 * hour)

    removed = render_com.sweep_stale_render_temp()

    assert set(removed) == {stale_src, stale_lo, stale_out}
    assert not os.path.isdir(stale_src)
    assert not os.path.isdir(stale_lo)
    assert not os.path.isdir(stale_out)
    assert os.path.isdir(fresh_src)
    assert os.path.isdir(fresh_out)
    assert os.path.isdir(unrelated)


def test_sweep_ignores_prefix_matching_files(tmp_path, monkeypatch):
    monkeypatch.setattr(render_com.tempfile, "gettempdir",
                        lambda: str(tmp_path))
    lone_file = tmp_path / "ppt_mcp_render_src_file"
    lone_file.write_bytes(b"not a dir")
    stamp = time.time() - 10 * 60 * 60
    os.utime(str(lone_file), (stamp, stamp))

    assert render_com.sweep_stale_render_temp() == []
    assert lone_file.is_file()


def test_render_slides_sweeps_before_running(tmp_path, monkeypatch):
    """render_slides must invoke the sweep on every call (pre-job)."""
    swept = []
    monkeypatch.setattr(render_com, "ensure_com_capability", lambda: None)
    monkeypatch.setattr(render_com, "sweep_stale_render_temp",
                        lambda: swept.append(True))
    monkeypatch.setattr(
        render_com._WORKER, "submit",
        lambda fn, timeout: {"paths": [], "width": 1280, "height": 720,
                             "slide_count": 0, "renderer": "powerpoint"})
    deck = _fake_deck(tmp_path)

    render_com.render_slides(deck, output_dir=str(tmp_path / "out"))

    assert swept == [True]
