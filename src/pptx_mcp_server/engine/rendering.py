"""
Rendering — convert PPTX slides to PNG images via LibreOffice headless.
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
import time

from .pptx_io import EngineError, ErrorCode

_CACHE_DIR = os.path.join(os.path.expanduser("~"), ".cache", "pptx-mcp-server", "renders")
_CACHE_MAX_AGE = 3600  # 1 hour in seconds


def _clean_old_renders() -> None:
    """Delete PNG files older than _CACHE_MAX_AGE from the cache directory."""
    if not os.path.isdir(_CACHE_DIR):
        return
    now = time.time()
    for fname in os.listdir(_CACHE_DIR):
        if not fname.endswith(".png"):
            continue
        fpath = os.path.join(_CACHE_DIR, fname)
        try:
            if now - os.path.getmtime(fpath) > _CACHE_MAX_AGE:
                os.remove(fpath)
        except OSError:
            pass


def _find_soffice() -> str:
    """Locate soffice binary (macOS, Linux, Windows)."""
    candidates = [shutil.which("soffice")]

    if platform.system() == "Darwin":
        candidates.extend([
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ])
    elif platform.system() == "Windows":
        import glob
        candidates.extend(
            glob.glob("C:/Program Files/LibreOffice*/program/soffice.exe")
            + glob.glob("C:/Program Files (x86)/LibreOffice*/program/soffice.exe")
        )
    else:  # Linux
        candidates.extend(["/usr/bin/soffice", "/usr/local/bin/soffice"])

    for c in candidates:
        if c and os.path.isfile(c):
            return c

    # Platform-specific install message
    if platform.system() == "Darwin":
        install_msg = "brew install --cask libreoffice"
    elif platform.system() == "Windows":
        install_msg = "Download from https://www.libreoffice.org/download/"
    else:
        install_msg = "sudo apt install libreoffice"

    raise EngineError(
        ErrorCode.INVALID_PARAMETER,
        f"LibreOffice (soffice) not found. Install: {install_msg}",
    )


def render_slide(
    file_path: str,
    slide_index: int = -1,
    output_dir: str = "",
    dpi: int = 150,
) -> str:
    """Render PPTX slide(s) to PNG images via LibreOffice + pdftoppm.

    Args:
        file_path: Path to .pptx file.
        slide_index: 0-based slide index to render. -1 = all slides.
        output_dir: Directory for output PNGs. Defaults to deterministic cache dir.
        dpi: Resolution (150 = good for review, 300 = print quality).

    Returns:
        Path(s) to rendered PNG file(s), one per line.
    """
    # Sanitize file_path to an absolute path
    file_path = os.path.abspath(file_path)

    if not os.path.exists(file_path):
        raise EngineError(ErrorCode.FILE_NOT_FOUND, f"File not found: {file_path}")

    soffice = _find_soffice()

    # Clean old cached renders
    _clean_old_renders()

    # Use a temp dir for intermediate PDF
    with tempfile.TemporaryDirectory() as tmpdir:
        # Step 1: PPTX → PDF via LibreOffice headless
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", file_path, "--outdir", tmpdir],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise EngineError(
                ErrorCode.INVALID_PPTX,
                f"LibreOffice conversion failed: {result.stderr}",
            )

        # Find the generated PDF
        basename = os.path.splitext(os.path.basename(file_path))[0]
        pdf_path = os.path.join(tmpdir, f"{basename}.pdf")
        if not os.path.exists(pdf_path):
            raise EngineError(
                ErrorCode.INVALID_PPTX,
                f"PDF not generated. LibreOffice output: {result.stdout}",
            )

        # Step 2: PDF → PNG
        if output_dir:
            out_dir = output_dir
            os.makedirs(output_dir, exist_ok=True)
        else:
            out_dir = _CACHE_DIR
            os.makedirs(_CACHE_DIR, exist_ok=True)

        page = slide_index + 1 if slide_index >= 0 else None

        pdftoppm_path = shutil.which("pdftoppm")
        if pdftoppm_path:
            # Fast path: pdftoppm (Unix/Mac)
            cmd = [pdftoppm_path, "-png", "-r", str(dpi)]
            if page is not None:
                cmd.extend(["-f", str(page), "-l", str(page)])
            cmd.extend([pdf_path, os.path.join(out_dir, "slide")])

            result2 = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            if result2.returncode != 0:
                raise EngineError(
                    ErrorCode.INVALID_PARAMETER,
                    f"pdftoppm failed: {result2.stderr}",
                )
        else:
            # Cross-platform fallback: pdf2image (Python library)
            try:
                from pdf2image import convert_from_path

                kwargs = {"dpi": dpi}
                if page is not None:
                    kwargs["first_page"] = page
                    kwargs["last_page"] = page
                images = convert_from_path(pdf_path, **kwargs)
                for i, img in enumerate(images):
                    pg = page if page is not None else i + 1
                    out_path = os.path.join(out_dir, f"slide-{pg:02d}.png")
                    img.save(out_path, "PNG")
            except ImportError:
                if platform.system() == "Darwin":
                    hint = "brew install poppler"
                elif platform.system() == "Windows":
                    hint = (
                        "Install poppler from "
                        "https://github.com/oschwartz10612/poppler-windows/releases/ "
                        "and add to PATH, or: pip install pdf2image"
                    )
                else:
                    hint = "sudo apt install poppler-utils"
                raise EngineError(
                    ErrorCode.INVALID_PARAMETER,
                    f"pdftoppm not found and pdf2image not installed. "
                    f"Install poppler: {hint}",
                )

        # Collect output PNGs
        pngs = sorted(
            [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.endswith(".png")]
        )

        if not pngs:
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                "No PNG files generated",
            )

        return "\n".join(pngs)


def render_slide_to_path(
    file_path: str,
    slide_index: int,
    output_path: str,
    dpi: int = 150,
) -> str:
    """Render a single slide to a specific output path.

    Convenience wrapper that renders one slide and moves
    the result to the desired path.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        result = render_slide(file_path, slide_index, output_dir=tmpdir, dpi=dpi)
        pngs = result.strip().split("\n")
        if pngs:
            shutil.move(pngs[0], output_path)
            return output_path
        raise EngineError(ErrorCode.INVALID_PARAMETER, "Render produced no output")
