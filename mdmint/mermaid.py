"""Mermaid diagram rendering — optional, with graceful fallback.

There is no pure-Python Mermaid renderer. If the mermaid CLI (``mmdc``, from
``@mermaid-js/mermaid-cli``) happens to be installed and on PATH, we shell out to
it and embed the resulting image. Otherwise we emit the diagram source as a styled
code block and warn once. ``mmdc`` is never required to install or use mdmint.
"""

from __future__ import annotations

import hashlib
import html
import logging
import os
import shutil
import struct
import subprocess
from pathlib import Path

log = logging.getLogger("mdmint")

_MMDC_PATH = None  # cached result of the PATH lookup (False once we know it's absent)
_WARNED = False

# mmdc emits PNGs at CSS-pixel dimensions; place them on the page treating those
# pixels as 96 DPI (xhtml2pdf's default image DPI).
_PNG_DPI = 96.0
# Cap diagram width to the printable page width so wide diagrams don't overflow the
# margin. A4 (210mm) minus the 2cm side margins set in resources/default.css.
_MAX_WIDTH_PT = (210.0 - 2 * 20.0) / 25.4 * 72.0  # ~481.9pt


def _find_mmdc() -> str | None:
    """Locate the mermaid CLI on PATH, caching the result for the process."""
    global _MMDC_PATH
    if _MMDC_PATH is None:
        # On Windows the executable is usually ``mmdc.cmd``; shutil.which handles
        # the extension resolution via PATHEXT.
        _MMDC_PATH = shutil.which("mmdc") or False
    return _MMDC_PATH or None


def is_snap_confined() -> bool:
    """True if the available mmdc is a snap package.

    Snap strict confinement cannot read ``/tmp`` or hidden dot-directories, so when
    rendering with a snap mmdc the work directory must live in a non-hidden path
    under the user's home (handled by :func:`mdmint.core._make_work_dir`).
    """
    mmdc = _find_mmdc()
    return bool(mmdc and "/snap/" in mmdc.replace("\\", "/"))


def _codeblock(source: str) -> str:
    """Render the diagram source as a styled, escaped code block (no logging)."""
    return f'<pre class="mermaid"><code>{html.escape(source)}</code></pre>'


def _png_size_pt(path: Path) -> tuple[float, float] | None:
    """Return a rendered PNG's on-page (width, height) in points, capped to the page
    width and preserving aspect ratio. ``None`` if the size can't be read.

    Reads the dimensions straight from the PNG IHDR header (no image library), so it
    works in the pure-Python core install.
    """
    try:
        with open(path, "rb") as fh:
            header = fh.read(24)
    except OSError:
        return None
    if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w_px, h_px = struct.unpack(">II", header[16:24])
    if not w_px or not h_px:
        return None
    w_pt = w_px * 72.0 / _PNG_DPI
    h_pt = h_px * 72.0 / _PNG_DPI
    if w_pt > _MAX_WIDTH_PT:  # scale down to fit; never upscale small diagrams
        h_pt *= _MAX_WIDTH_PT / w_pt
        w_pt = _MAX_WIDTH_PT
    return w_pt, h_pt


def _warn_once(message: str) -> None:
    global _WARNED
    if not _WARNED:
        log.warning(message)
        _WARNED = True


def render(source: str, work_dir: Path) -> str:
    """Return an HTML snippet for a single mermaid diagram.

    If ``mmdc`` is available, render ``source`` to a PNG inside ``work_dir`` and
    return an ``<img>`` tag pointing at the absolute path. Otherwise fall back to a
    code block. Rendering failures also fall back, so a bad diagram never aborts the
    whole conversion.
    """
    mmdc = _find_mmdc()
    if not mmdc:
        _warn_once(
            "mermaid: no local renderer found (mmdc not on PATH); rendering "
            "diagrams as code blocks. Install '@mermaid-js/mermaid-cli' for real "
            "diagrams."
        )
        return _codeblock(source)

    digest = hashlib.sha1(source.encode("utf-8")).hexdigest()[:16]
    out_path = work_dir / f"mermaid-{digest}.png"

    if not out_path.exists():
        in_path = work_dir / f"mermaid-{digest}.mmd"
        in_path.write_text(source, encoding="utf-8")
        cmd = [
            mmdc,
            "-i",
            str(in_path),
            "-o",
            str(out_path),
            "-b",
            "transparent",
        ]
        try:
            subprocess.run(
                cmd,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=60,
            )
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as exc:
            detail = getattr(exc, "stderr", b"")
            if isinstance(detail, bytes):
                detail = detail.decode("utf-8", "replace")
            _warn_once(
                "mermaid: rendering failed (%s); falling back to code block. %s"
                % (type(exc).__name__, detail.strip())
            )
            return _codeblock(source)

    src = os.path.abspath(out_path)
    size = _png_size_pt(out_path)
    style = f' style="width:{size[0]:.1f}pt;height:{size[1]:.1f}pt"' if size else ""
    return (
        f'<p class="mermaid-diagram"><img src="{html.escape(src)}"{style} '
        'alt="mermaid diagram" /></p>'
    )
