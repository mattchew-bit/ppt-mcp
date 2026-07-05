"""LibreOffice fallback renderer (approximate fidelity, any platform).

Strictly optional and auto-selected only when PowerPoint COM rendering is
unavailable. Pipeline: ``soffice --headless --convert-to pdf`` -> rasterize
pages via pypdfium2 (Apache/BSD; PyMuPDF is AGPL and prohibited here).

Notes locked by the research findings:
* never ``--convert-to png`` -- for Impress documents it exports slide 1 only
* throwaway ``-env:UserInstallation`` profile per run (no profile-lock
  collisions with a user's running LibreOffice)
* output PNGs are tagged ``renderer: libreoffice`` -- ``compare_renders``
  refuses to pixel-compare them against PowerPoint output
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional

from utils.render_com import (
    DEFAULT_RENDER_WIDTH,
    RenderCapabilityError,
    validate_render_source,
    validate_render_width,
)
from utils.resolve_analysis import parse_slide_range

RENDERER_LIBREOFFICE = "libreoffice"
DEFAULT_LO_TIMEOUT_S = 300.0

_WINDOWS_SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)


def find_soffice() -> Optional[str]:
    """Absolute path of the ``soffice`` binary, or ``None`` when absent."""
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return os.path.abspath(found)
    for candidate in _WINDOWS_SOFFICE_CANDIDATES:
        if os.path.isfile(candidate):
            return candidate
    return None


def ensure_lo_capability() -> str:
    """Return the soffice path or raise an actionable capability error."""
    soffice = find_soffice()
    if soffice is None:
        raise RenderCapabilityError(
            "LibreOffice fallback renderer unavailable: 'soffice' was not "
            "found on PATH or in the default install locations. Install "
            "LibreOffice (https://www.libreoffice.org/) for approximate "
            "rendering, or use Windows with desktop PowerPoint plus "
            "pip install 'ppt-mcp[render]' for full-fidelity rendering."
        )
    return soffice


def _ensure_pypdfium2():
    try:
        import pypdfium2
        return pypdfium2
    except ImportError as exc:
        raise RenderCapabilityError(
            "The LibreOffice fallback renderer requires pypdfium2 to "
            "rasterize the intermediate PDF. Install the render extra: "
            "pip install 'ppt-mcp[render]'."
        ) from exc


def _convert_to_pdf(soffice: str, abs_source: str, work_dir: str,
                    timeout: float) -> str:
    """Run soffice headless with a throwaway profile; return the PDF path."""
    profile_dir = os.path.join(work_dir, "profile")
    os.makedirs(profile_dir, exist_ok=True)
    profile_url = "file:///" + profile_dir.replace("\\", "/").lstrip("/")
    command = [
        soffice,
        "--headless", "--norestore", "--nolockcheck",
        f"-env:UserInstallation={profile_url}",
        "--convert-to", "pdf",
        "--outdir", work_dir,
        abs_source,
    ]
    try:
        completed = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"LibreOffice PDF conversion timed out after {timeout:.0f}s "
            f"for '{abs_source}'."
        ) from exc

    pdf_path = os.path.join(work_dir, Path(abs_source).stem + ".pdf")
    if completed.returncode != 0 or not os.path.isfile(pdf_path):
        detail = (completed.stderr or completed.stdout or "").strip()
        raise RuntimeError(
            f"LibreOffice failed to convert '{abs_source}' to PDF "
            f"(exit code {completed.returncode})."
            + (f" Output: {detail[:500]}" if detail else "")
        )
    return pdf_path


def _rasterize_pdf(pdf_path: str, indices, width: int, height: int,
                   out_dir: str, stem: str) -> list:
    """Rasterize selected PDF pages to exact-width PNGs via pypdfium2."""
    pypdfium2 = _ensure_pypdfium2()
    paths = []
    pdf = pypdfium2.PdfDocument(pdf_path)
    try:
        for idx in indices:
            page = pdf[idx]
            scale = width / float(page.get_width())
            pil_image = page.render(scale=scale).to_pil().convert("RGB")
            if pil_image.size != (width, height):
                pil_image = pil_image.resize((width, height))
            out_path = os.path.abspath(
                os.path.join(out_dir, f"{stem}_slide_{idx + 1:03d}.png"))
            pil_image.save(out_path, "PNG")
            paths.append(out_path)
    finally:
        pdf.close()
    return paths


def render_slides_lo(file_path: str,
                     width: int = DEFAULT_RENDER_WIDTH,
                     slide_range: Optional[str] = None,
                     output_dir: Optional[str] = None,
                     timeout: float = DEFAULT_LO_TIMEOUT_S) -> Dict[str, Any]:
    """Render deck slides to PNG via LibreOffice (approximate fidelity).

    Same contract as ``utils.render_com.render_slides`` but with
    ``renderer: "libreoffice"`` -- fonts may substitute, SmartArt goes
    static, charts re-render. Never pixel-compare against PowerPoint
    output (``compare_renders`` enforces this via the PNG renderer tag).
    """
    soffice = ensure_lo_capability()
    _ensure_pypdfium2()  # fail fast before spending time on conversion
    abs_source = validate_render_source(file_path)
    validate_render_width(width)

    from utils.render_com import (
        TEMP_LO_PREFIX,
        _prepare_output_dir,
        sweep_stale_render_temp,
    )
    from utils.render_compare import tag_png_renderer

    sweep_stale_render_temp()
    out_dir = _prepare_output_dir(output_dir, abs_source)
    work_dir = tempfile.mkdtemp(prefix=TEMP_LO_PREFIX)
    try:
        pdf_path = _convert_to_pdf(soffice, abs_source, work_dir, timeout)
        pypdfium2 = _ensure_pypdfium2()
        pdf = pypdfium2.PdfDocument(pdf_path)
        try:
            slide_count = len(pdf)
            indices = parse_slide_range(slide_range, slide_count)
            first_page = pdf[0]
            height = max(1, round(
                width * float(first_page.get_height())
                / float(first_page.get_width())))
        finally:
            pdf.close()

        paths = _rasterize_pdf(
            pdf_path, indices, width, height, out_dir, Path(abs_source).stem)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    for png_path in paths:
        tag_png_renderer(png_path, RENDERER_LIBREOFFICE)
    return {
        "paths": paths,
        "width": width,
        "height": height,
        "slide_count": slide_count,
        "renderer": RENDERER_LIBREOFFICE,
    }
