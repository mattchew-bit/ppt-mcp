"""PowerPoint COM slide rendering (Windows + desktop PowerPoint).

Design (locked by findings/style-fidelity-research-2026-07-03.md section 2):

* **Dedicated STA worker thread + job queue** -- MCP handlers run on a
  threadpool; ALL COM calls funnel through one thread that calls
  ``pythoncom.CoInitialize()`` once. Jobs carry a per-job timeout; on
  timeout the caller triggers emergency cleanup (kill POWERPNT only when
  WE launched the process and no foreign presentations were open in it).
* **Ownership tracking** -- ``GetActiveObject`` attaches to a running
  PowerPoint when present; only when that fails do we ``Dispatch`` (launch)
  and record the fact plus the process id. ``app.Quit()`` happens ONLY if
  we launched the process AND ``Presentations.Count == 0`` after our work.
  Presentations we did not open are never closed. ``app.Visible`` is never
  set to False (PowerPoint raises -2147188160).
* **Macro safety** -- prior ``app.AutomationSecurity`` is saved, forced to
  ``msoAutomationSecurityForceDisable`` (3) before any Open, and restored
  in ``finally``. Decks open ``ReadOnly=True, Untitled=False,
  WithWindow=False``; ``DisplayAlerts`` is set to ``ppAlertsNone`` (1).
* **Temp-copy rendering** -- the deck is copied to a temp path and the copy
  is rendered, so a user having the original open in PowerPoint never
  blocks a render. The copy is deleted afterwards.
* **Container pre-check** -- OLE magic bytes (``D0 CF 11 E0`` = encrypted /
  legacy binary) are refused BEFORE any COM call (a blocking password
  dialog otherwise); empty and non-ZIP files are refused too (PowerPoint
  silently opens a zero-byte file as a blank presentation).
* **Stale temp sweep** -- every render first reaps stale
  ``ppt_mcp_render_src_*`` / ``ppt_mcp_lo_*`` work dirs (>1h) and
  ``ppt_mcp_render_*`` default output dirs (>7d) from %TEMP%, bounding
  accumulation across crashed or killed sessions.
* **Explicit pixel dims** -- ``Slide.Export(abs_path, "PNG", w, h)`` with
  ``h`` computed from the ``PageSetup`` slide ratio; absolute paths only
  (COM resolves relative paths against PowerPoint's CWD). Whole-deck
  export loops per slide -- never ``SaveAs ppSaveAsPNG`` (registry-DPI,
  localized filenames).
* **Lazy imports** -- pywin32 is imported inside functions so this module
  (and the MCP server) imports cleanly on machines without pywin32;
  calling render entry points there raises :class:`RenderCapabilityError`.
"""

from __future__ import annotations

import contextlib
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set

from utils.resolve_analysis import parse_slide_range

# ---------------------------------------------------------------- constants

OLE_MAGIC = b"\xd0\xcf\x11\xe0"
ZIP_MAGIC = b"PK\x03\x04"
#: Smallest possible ZIP container (an empty archive's end-of-central-
#: directory record is 22 bytes); anything smaller is truncated garbage.
MIN_ZIP_BYTES = 22

DEFAULT_RENDER_WIDTH = 1280
MIN_RENDER_WIDTH = 32
MAX_RENDER_WIDTH = 4096
DEFAULT_JOB_TIMEOUT_S = 180.0

RENDERER_POWERPOINT = "powerpoint"

RENDERABLE_EXTENSIONS = (".pptx", ".pptm", ".ppsx", ".potx")

# Temp-directory families created by the renderers (all under %TEMP%):
#   ppt_mcp_render_src_*   -- per-job deck copies (deleted by the job; a
#                             crash or a kill-locked copy can strand one)
#   ppt_mcp_lo_*           -- LibreOffice work dirs (same lifecycle)
#   ppt_mcp_render_<stem>_* -- DEFAULT output dirs, returned to the caller
#                             and intentionally left on disk
# ``sweep_stale_render_temp`` reaps stale members of all three families.
TEMP_SRC_PREFIX = "ppt_mcp_render_src_"
TEMP_LO_PREFIX = "ppt_mcp_lo_"
TEMP_OUT_PREFIX = "ppt_mcp_render_"
#: Work dirs live only for the duration of one job (<= the job timeout);
#: one hour of slack means a sweep can never race a live job.
STALE_WORK_DIR_MAX_AGE_S = 60.0 * 60.0
#: Default output dirs are the caller's artifacts; keep them a week.
STALE_OUTPUT_DIR_MAX_AGE_S = 7.0 * 24.0 * 60.0 * 60.0

# Office / PowerPoint enum values (raw ints; no makepy dependency)
MSO_TRUE = -1
MSO_FALSE = 0
PP_ALERTS_NONE = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class RenderCapabilityError(RuntimeError):
    """Rendering is unavailable on this machine (missing dep / platform)."""


class RenderTimeoutError(RuntimeError):
    """A render job exceeded its timeout on the COM worker thread."""


# ------------------------------------------------------- boundary validation


def check_not_encrypted(file_path: str) -> None:
    """Refuse encrypted / legacy / corrupt containers BEFORE any COM call.

    Password-protected .pptx files are OLE containers (magic ``D0 CF 11
    E0``), and opening one via COM blocks on a password dialog. Legacy
    binary .ppt files share the magic and are equally unrenderable here.
    Everything renderable is a ZIP (OOXML) container, so a file without
    the ``PK`` zip signature -- including a zero-byte file, which
    PowerPoint silently opens as a blank 0-slide presentation -- is
    refused here with an actionable message instead of reaching COM.
    """
    abs_path = os.path.abspath(file_path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"Presentation not found: {abs_path}")
    size = os.path.getsize(abs_path)
    with open(abs_path, "rb") as handle:
        header = handle.read(8)
    if header.startswith(OLE_MAGIC):
        raise ValueError(
            f"'{abs_path}' is an OLE compound file -- most likely a "
            "password-protected (encrypted) presentation, or a legacy "
            "binary .ppt. Rendering requires an unencrypted OOXML .pptx: "
            "remove the password (File > Info > Protect Presentation) or "
            "re-save as .pptx, then retry."
        )
    if size == 0:
        raise ValueError(
            f"'{abs_path}' is empty (0 bytes) -- there is nothing to "
            "render. Re-export or re-save the presentation, then retry."
        )
    if size < MIN_ZIP_BYTES or not header.startswith(ZIP_MAGIC):
        raise ValueError(
            f"'{abs_path}' is not a ZIP-based OOXML presentation (missing "
            f"the 'PK' zip signature, file size {size} bytes). The file is "
            "most likely corrupt, truncated, or not a real PowerPoint "
            "deck despite its extension. Re-save it from PowerPoint and "
            "retry."
        )


def validate_render_source(file_path: str) -> str:
    """Validate a deck path for rendering; returns the absolute path."""
    abs_path = os.path.abspath(str(file_path))
    extension = os.path.splitext(abs_path)[1].lower()
    if extension not in RENDERABLE_EXTENSIONS:
        raise ValueError(
            f"Unsupported presentation extension '{extension}' for "
            f"'{abs_path}'. Renderable formats: "
            f"{', '.join(RENDERABLE_EXTENSIONS)}"
        )
    check_not_encrypted(abs_path)
    return abs_path


def validate_render_width(width: int) -> int:
    """Validate the requested pixel width at the tool boundary."""
    if not isinstance(width, int) or isinstance(width, bool):
        raise ValueError(f"width must be an integer, got {width!r}")
    if width < MIN_RENDER_WIDTH or width > MAX_RENDER_WIDTH:
        raise ValueError(
            f"width {width} out of range: must be between "
            f"{MIN_RENDER_WIDTH} and {MAX_RENDER_WIDTH} pixels"
        )
    return width


def ensure_com_capability() -> None:
    """Raise a capability error unless PowerPoint COM rendering can work."""
    if sys.platform != "win32":
        raise RenderCapabilityError(
            "PowerPoint COM rendering requires Windows with desktop "
            "PowerPoint installed (current platform: "
            f"{sys.platform}). On other platforms, install LibreOffice "
            "for the approximate fallback renderer."
        )
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise RenderCapabilityError(
            "PowerPoint COM rendering requires pywin32, which is not "
            "installed. Install the render extra: "
            "pip install 'ppt-mcp[render]'. Rendering also requires "
            "Windows with desktop PowerPoint installed."
        ) from exc
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                            "PowerPoint.Application"):
            pass
    except OSError as exc:
        raise RenderCapabilityError(
            "PowerPoint COM rendering requires desktop PowerPoint, but "
            "'PowerPoint.Application' is not registered on this machine. "
            "Install desktop PowerPoint (Microsoft 365 / Office), or "
            "install LibreOffice for the approximate fallback renderer."
        ) from exc


# --------------------------------------------------------- STA worker thread


@contextlib.contextmanager
def _faulthandler_paused():
    """Silence benign first-chance SEH noise during COM jobs.

    Releasing the last proxy reference to a PowerPoint we just ``Quit()``
    raises RPC_E_DISCONNECTED (0x80010108) as a first-chance structured
    exception. COM handles it internally, but ``faulthandler`` (enabled by
    pytest) prints a scary "Windows fatal exception" dump for it. Pausing
    faulthandler for the duration of a render job suppresses only that
    noise; a genuine hard crash would still take the process down visibly.
    """
    import faulthandler

    was_enabled = faulthandler.is_enabled()
    if was_enabled:
        faulthandler.disable()
    try:
        yield
    finally:
        if was_enabled:
            faulthandler.enable()


class _Job:
    """One unit of work executed on the COM worker thread.

    ``started`` fires when the worker dequeues the job, so a caller's
    timeout budget covers only the job's own run time -- never time spent
    queued behind another render. ``cancelled`` is set by a caller that
    gave up waiting in the queue; the worker then skips the job entirely.
    """

    __slots__ = ("fn", "started", "cancelled", "done", "result", "error")

    def __init__(self, fn: Callable[[], Any]):
        self.fn = fn
        self.started = threading.Event()
        self.cancelled = threading.Event()
        self.done = threading.Event()
        self.result: Any = None
        self.error: Optional[BaseException] = None


class _ComWorker:
    """Serializes ALL COM access onto one CoInitialize'd STA thread.

    PowerPoint is single-instance per machine, so renders are serialized
    by construction: one queue, one thread, one job at a time.
    """

    def __init__(self):
        self._queue: "queue.SimpleQueue[_Job]" = queue.SimpleQueue()
        self._thread: Optional[threading.Thread] = None
        self._start_lock = threading.Lock()
        self._state_lock = threading.Lock()
        # Ownership state, written by jobs on the worker thread and read
        # by emergency cleanup on the calling thread.
        self._launched_pid: Optional[int] = None
        self._foreign_presentations = False

    # -- lifecycle ---------------------------------------------------------

    def _ensure_thread(self) -> None:
        with self._start_lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._thread = threading.Thread(
                target=self._main, name="ppt-mcp-com-render", daemon=True)
            self._thread.start()

    def _main(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        try:
            self._run_jobs()
        finally:  # pragma: no cover - daemon thread teardown
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    def _run_jobs(self) -> None:
        """Job loop, separated from COM init so tests can run it directly."""
        while True:
            job = self._queue.get()
            if job.cancelled.is_set():  # caller gave up while queued
                job.done.set()
                continue
            job.started.set()
            try:
                with _faulthandler_paused():
                    job.result = job.fn()
            except BaseException as exc:  # propagate to the caller
                job.error = exc
            finally:
                job.done.set()

    # -- ownership bookkeeping (called from jobs on the worker thread) ------

    def record_launch(self, pid: Optional[int]) -> None:
        with self._state_lock:
            self._launched_pid = pid

    def record_quit(self) -> None:
        with self._state_lock:
            self._launched_pid = None

    def record_foreign_presentations(self, present: bool) -> None:
        with self._state_lock:
            self._foreign_presentations = present

    # -- submission ----------------------------------------------------------

    def submit(self, fn: Callable[[], Any], timeout: float) -> Any:
        """Run ``fn`` on the worker; ``timeout`` bounds each phase.

        The budget is applied twice, independently: once for waiting in
        the queue (another render may be running -- a queue timeout kills
        NOTHING, since the running job is someone else's healthy work) and
        once
        for the job's own run time (a run timeout triggers emergency
        cleanup of a genuinely hung PowerPoint).
        """
        self._ensure_thread()
        job = _Job(fn)
        self._queue.put(job)
        if not job.started.wait(timeout):
            job.cancelled.set()
            raise RenderTimeoutError(
                f"Render job timed out after {timeout:.0f}s waiting for "
                "the render worker -- another render job is still "
                "running. Nothing was terminated; retry once the current "
                "render finishes."
            )
        if not job.done.wait(timeout):
            killed = self._emergency_cleanup()
            raise RenderTimeoutError(
                f"Render job timed out after {timeout:.0f}s. "
                + ("The PowerPoint process this server launched was "
                   "terminated." if killed else
                   "PowerPoint was left running: either this server did "
                   "not launch it, other presentations were open in it, "
                   "or it could not be terminated. Close any dialog in "
                   "PowerPoint (check Task Manager for POWERPNT.EXE) and "
                   "retry.")
            )
        if job.error is not None:
            raise job.error
        return job.result

    def _emergency_cleanup(self) -> bool:
        """Kill POWERPNT after a hung job -- only when it is safe.

        Safe means: WE launched the process, it is still a POWERPNT.EXE
        process right now (a recorded pid can go stale and be reused by
        an unrelated process once Windows recycles it), AND, as of the
        last job start, it contained no presentations we did not open. A
        user's own PowerPoint session is one we attached to (never
        launched) and is never killed; a user who attached to OUR
        instance mid-job is protected by the foreign-presentations flag.

        Returns True only when taskkill actually succeeded, so the
        timeout message never claims a kill that did not happen; on kill
        failure the pid is kept for a retry on the next timeout.
        """
        with self._state_lock:
            pid = self._launched_pid
            foreign = self._foreign_presentations
        if pid is None or foreign:
            return False
        if pid not in _powerpnt_pids():
            # Stale record: the process already exited (or the pid was
            # reused by something that is not PowerPoint). Forget it so a
            # recycled pid can never be taskkilled by a later timeout.
            self.record_quit()
            return False
        completed = subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            capture_output=True,
        )
        if completed.returncode != 0:
            return False
        self.record_quit()
        return True


_WORKER = _ComWorker()


# ------------------------------------------------------------ COM job pieces


def _powerpnt_pids() -> Set[int]:
    """PIDs of every POWERPNT.EXE currently running (empty set on failure)."""
    try:
        completed = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE",
             "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=30,
        )
    except (OSError, subprocess.SubprocessError):
        return set()
    pids: Set[int] = set()
    for line in completed.stdout.splitlines():
        columns = [col.strip('"') for col in line.strip().split('","')]
        if len(columns) < 2 or columns[0].upper() != "POWERPNT.EXE":
            continue
        with contextlib.suppress(ValueError):
            pids.add(int(columns[1]))
    return pids


def _attach_or_launch():
    """Return ``(app, we_launched, launched_pid)``; attach first.

    The launched PID is derived by diffing the POWERPNT process list
    around the ``Dispatch`` call. It can NOT be read from ``app.HWND``:
    on late-bound Dispatch against Office16 that property raises
    ``com_error`` DISP_E_MEMBERNOTFOUND, which used to leave the
    emergency kill switch permanently disarmed (``_app_pid`` remains as
    a secondary path for early-bound setups where HWND does resolve).
    """
    import pythoncom
    import win32com.client

    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        return app, False, None
    except pythoncom.com_error:
        pass
    pids_before = _powerpnt_pids()
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
    except pythoncom.com_error as exc:
        raise RenderCapabilityError(
            "Could not start PowerPoint via COM -- is desktop PowerPoint "
            f"installed on this machine? (COM error: {exc})"
        ) from exc
    return app, True, _launched_pid(pids_before, app)


def _launched_pid(pids_before: Set[int], app) -> Optional[int]:
    """PID of the POWERPNT we just launched, or None when ambiguous."""
    new_pids = _powerpnt_pids() - pids_before
    if len(new_pids) == 1:
        return new_pids.pop()
    # 0 or 2+ new processes (racing launches): fall back to the HWND
    # route; a None result just means the kill switch stays disarmed.
    return _app_pid(app)


def _app_pid(app) -> Optional[int]:
    """Process id via ``app.HWND`` -- works only on early-bound proxies."""
    try:
        import win32process

        _, pid = win32process.GetWindowThreadProcessId(app.HWND)
        return int(pid) or None
    except Exception:
        return None


def _quit_if_owned_and_empty(app, we_launched: bool) -> None:
    """``app.Quit()`` only if we launched it and nothing remains open."""
    if not we_launched:
        return
    with contextlib.suppress(Exception):
        if app.Presentations.Count == 0:
            app.Quit()
            _WORKER.record_quit()


def _export_pngs(pres, slide_range: Optional[str], width: int,
                 out_dir: str, stem: str) -> Dict[str, Any]:
    """Export the selected slides of an open presentation as PNGs."""
    slide_count = int(pres.Slides.Count)
    indices = parse_slide_range(slide_range, slide_count)

    page = pres.PageSetup
    height = max(1, round(
        width * float(page.SlideHeight) / float(page.SlideWidth)))

    paths: List[str] = []
    for idx in indices:
        out_path = os.path.abspath(
            os.path.join(out_dir, f"{stem}_slide_{idx + 1:03d}.png"))
        pres.Slides(idx + 1).Export(out_path, "PNG", width, height)
        paths.append(out_path)
    return {
        "paths": paths,
        "width": width,
        "height": height,
        "slide_count": slide_count,
        "renderer": RENDERER_POWERPOINT,
    }


def _cleanup_job(app, pres, we_launched: bool,
                 saved_security, saved_alerts) -> None:
    """Best-effort teardown: close OUR presentation, restore app state."""
    if pres is not None:
        with contextlib.suppress(Exception):
            pres.Saved = MSO_TRUE  # suppress any save prompt
        with contextlib.suppress(Exception):
            pres.Close()
    if saved_security is not None:
        with contextlib.suppress(Exception):
            app.AutomationSecurity = saved_security
    if saved_alerts is not None:
        with contextlib.suppress(Exception):
            app.DisplayAlerts = saved_alerts
    _quit_if_owned_and_empty(app, we_launched)


def _open_presentation(app, temp_copy: str, abs_source: str):
    """Open the temp copy; map COM open failures to an actionable error."""
    import pythoncom

    try:
        return app.Presentations.Open(
            os.path.abspath(temp_copy),
            MSO_TRUE,   # ReadOnly
            MSO_FALSE,  # Untitled
            MSO_FALSE,  # WithWindow
        )
    except pythoncom.com_error as exc:
        raise ValueError(
            f"PowerPoint could not open '{abs_source}' -- the file is "
            "most likely corrupt, truncated, or not a real PowerPoint "
            "presentation despite its extension. Re-save it from "
            f"PowerPoint and retry. (COM detail: {exc})"
        ) from exc


def _run_com_export(abs_source: str, temp_copy: str,
                    slide_range: Optional[str], width: int,
                    out_dir: str) -> Dict[str, Any]:
    """Attach/launch PowerPoint, open the temp copy, export the PNGs."""
    app, we_launched, launched_pid = _attach_or_launch()
    if we_launched:
        _WORKER.record_launch(launched_pid)
    saved_alerts = None
    saved_security = None
    pres = None
    try:
        with contextlib.suppress(Exception):
            _WORKER.record_foreign_presentations(app.Presentations.Count > 0)
        with contextlib.suppress(Exception):
            saved_alerts = app.DisplayAlerts
            app.DisplayAlerts = PP_ALERTS_NONE
        saved_security = app.AutomationSecurity
        app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE

        pres = _open_presentation(app, temp_copy, abs_source)
        return _export_pngs(pres, slide_range, width, out_dir,
                            Path(abs_source).stem)
    finally:
        _cleanup_job(app, pres, we_launched, saved_security, saved_alerts)


def _render_job(abs_source: str, slide_range: Optional[str], width: int,
                out_dir: str) -> Dict[str, Any]:
    """Runs ON the worker thread: open a temp copy, export PNGs, clean up.

    The temp-dir removal wraps EVERYTHING -- including the deck copy and
    the attach/launch -- so no ``ppt_mcp_render_src_*`` deck copy is left
    in %TEMP% on any failure path (deck copies are confidential client
    content). A copy still locked by a just-killed POWERPNT survives the
    ``rmtree``; ``sweep_stale_render_temp`` reaps it on a later call.
    """
    temp_dir = tempfile.mkdtemp(prefix=TEMP_SRC_PREFIX)
    try:
        temp_copy = os.path.join(temp_dir, os.path.basename(abs_source))
        shutil.copy2(abs_source, temp_copy)
        return _run_com_export(abs_source, temp_copy,
                               slide_range, width, out_dir)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------- public API


def _stale_max_age(dir_name: str) -> Optional[float]:
    """Max age for a render temp dir, or None when it is not one of ours."""
    if dir_name.startswith((TEMP_SRC_PREFIX, TEMP_LO_PREFIX)):
        return STALE_WORK_DIR_MAX_AGE_S
    if dir_name.startswith(TEMP_OUT_PREFIX):
        return STALE_OUTPUT_DIR_MAX_AGE_S
    return None


def sweep_stale_render_temp(now: Optional[float] = None) -> List[str]:
    """Best-effort reap of stale render temp dirs from earlier sessions.

    Windows never purges %TEMP%, so anything stranded there -- deck
    copies orphaned by a process crash or a kill-locked file, and
    default output dirs nobody deleted -- accumulates forever without
    this. Runs before each render; failures are skipped silently (a
    locked dir just waits for the next sweep). Returns removed paths.
    """
    removed: List[str] = []
    reference = time.time() if now is None else now
    temp_root = tempfile.gettempdir()
    try:
        names = os.listdir(temp_root)
    except OSError:
        return removed
    for name in names:
        max_age = _stale_max_age(name)
        if max_age is None:
            continue
        path = os.path.join(temp_root, name)
        with contextlib.suppress(OSError):
            if not os.path.isdir(path):
                continue
            if reference - os.path.getmtime(path) < max_age:
                continue
            shutil.rmtree(path, ignore_errors=True)
            if not os.path.isdir(path):
                removed.append(path)
    return removed


def _prepare_output_dir(output_dir: Optional[str], abs_source: str) -> str:
    """Resolve/create the PNG output dir.

    With no explicit ``output_dir`` a fresh ``ppt_mcp_render_<stem>_*``
    dir is created under %TEMP%; it is returned to the caller and left
    on disk (the PNGs ARE the result), then reaped by
    ``sweep_stale_render_temp`` once older than a week.
    """
    if output_dir:
        resolved = os.path.abspath(output_dir)
        os.makedirs(resolved, exist_ok=True)
        return resolved
    stem = Path(abs_source).stem
    return tempfile.mkdtemp(prefix=f"{TEMP_OUT_PREFIX}{stem}_")


def render_slides(file_path: str,
                  width: int = DEFAULT_RENDER_WIDTH,
                  slide_range: Optional[str] = None,
                  output_dir: Optional[str] = None,
                  timeout: float = DEFAULT_JOB_TIMEOUT_S) -> Dict[str, Any]:
    """Render slides of a deck to PNG via PowerPoint COM.

    Args:
        file_path: Path to the .pptx deck (rendered from a temp copy).
        width: Output pixel width; height follows the slide aspect ratio.
        slide_range: 1-based selection like ``"1-3,5"``; ``None`` = all.
        output_dir: Directory for PNGs (default: a fresh temp directory).
        timeout: Timeout in seconds, applied separately to waiting for
            the render worker (behind other renders) and to the job's
            own run time on it.

    Returns:
        ``{"paths": [...], "width": w, "height": h, "slide_count": n,
        "renderer": "powerpoint"}``
    """
    ensure_com_capability()
    abs_source = validate_render_source(file_path)
    validate_render_width(width)
    sweep_stale_render_temp()
    out_dir = _prepare_output_dir(output_dir, abs_source)

    result = _WORKER.submit(
        lambda: _render_job(abs_source, slide_range, width, out_dir),
        timeout=timeout,
    )

    from utils.render_compare import tag_png_renderer

    for png_path in result["paths"]:
        tag_png_renderer(png_path, RENDERER_POWERPOINT)
    return result
