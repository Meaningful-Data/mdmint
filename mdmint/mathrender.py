"""LaTeX math rendering — optional, with graceful fallback.

Static output formats (PDF/DOCX) cannot run MathJax, so math has to become an
image. If ``matplotlib`` is installed we render each expression with its built-in
``mathtext`` engine (a sizeable LaTeX-math subset, no system LaTeX required) to a
transparent PNG. If matplotlib is absent, or an expression isn't supported, we fall
back to the literal source text and warn once.
"""

from __future__ import annotations

import hashlib
import html
import logging
import os
from pathlib import Path
from types import ModuleType
from typing import Literal

log = logging.getLogger("mdmint")

_DPI = 200  # render resolution; combined with pixel size to derive on-page pt size
_MATPLOTLIB: ModuleType | Literal[False] | None = None  # None = not probed yet
_WARNED = False
_cache: dict[str, str] = {}


def _get_matplotlib() -> ModuleType | None:
    global _MATPLOTLIB
    if _MATPLOTLIB is None:
        try:
            import matplotlib

            matplotlib.use("Agg")  # headless backend, no display needed
            _MATPLOTLIB = matplotlib
        except Exception:  # pragma: no cover - depends on environment
            _MATPLOTLIB = False
    return _MATPLOTLIB or None


def _warn_missing() -> None:
    global _WARNED
    if not _WARNED:
        log.warning(
            "math: matplotlib not installed; rendering math as literal text. "
            "Install with 'pip install mdmint[math]' for rendered equations."
        )
        _WARNED = True


def _fallback(expr: str, block: bool) -> str:
    escaped = html.escape(expr)
    if block:
        return f'<p class="math-block-fallback"><code>$$ {escaped} $$</code></p>'
    return f'<code class="math-inline-fallback">${escaped}$</code>'


def render(expr: str, block: bool, work_dir: Path) -> str:
    """Return an HTML snippet (an ``<img>`` or literal fallback) for one expression."""
    mpl = _get_matplotlib()
    if mpl is None:
        _warn_missing()
        return _fallback(expr, block)

    key = ("B:" if block else "I:") + expr
    if key in _cache:
        return _cache[key]

    digest = hashlib.sha1(key.encode("utf-8"), usedforsecurity=False).hexdigest()[:16]
    out_path = work_dir / f"math-{digest}.png"

    try:
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        from matplotlib.figure import Figure

        fig = Figure(figsize=(0.1, 0.1))
        FigureCanvasAgg(fig)
        # mathtext expects the expression wrapped in $...$.
        fig.text(0, 0, f"${expr}$", fontsize=14)
        fig.savefig(
            out_path,
            dpi=_DPI,
            format="png",
            bbox_inches="tight",
            pad_inches=0.02,
            transparent=True,
        )

        # Derive on-page size in points so the image matches surrounding text.
        from PIL import Image  # Pillow ships as a matplotlib dependency

        with Image.open(out_path) as im:
            w_px, h_px = im.size
        h_pt = h_px / _DPI * 72.0
        w_pt = w_px / _DPI * 72.0
    except Exception as exc:
        log.warning("math: could not render %r (%s); using literal text.", expr, type(exc).__name__)
        result = _fallback(expr, block)
        _cache[key] = result
        return result

    src = html.escape(os.path.abspath(out_path))
    if block:
        snippet = (
            f'<p class="math-block"><img class="math-block" src="{src}" '
            f'alt="{html.escape(expr)}" style="height:{h_pt:.1f}pt" /></p>'
        )
    else:
        snippet = (
            f'<img class="math-inline" src="{src}" alt="{html.escape(expr)}" '
            f'style="height:{h_pt:.1f}pt;width:{w_pt:.1f}pt" />'
        )
    _cache[key] = snippet
    return snippet
