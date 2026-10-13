"""Image asset resolution.

After a Markdown body is rendered to HTML, every ``<img src>`` is rewritten to an
absolute local filesystem path so that both writers (htmldocx and the PDF backends)
resolve images identically:

* relative paths   → resolved against the source ``.md`` file's directory
* ``http(s)`` URLs → downloaded into the work directory (offline-safe thereafter)
* absolute paths   → kept (these include mermaid/math images we generated)
* ``data:`` URIs   → left untouched

Anything that cannot be resolved is left as-is with a warning rather than aborting.
"""

from __future__ import annotations

import hashlib
import logging
import mimetypes
import os
import re
import urllib.request
from pathlib import Path

log = logging.getLogger("mdmint")

# Matches the src attribute of an <img> tag, capturing the quote and the value.
_IMG_SRC = re.compile(r'(<img\b[^>]*?\ssrc=)(["\'])(.*?)\2', re.IGNORECASE)

_EXT_BY_TYPE = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/svg+xml": ".svg",
    "image/webp": ".webp",
}


def _download(url: str, work_dir: Path) -> str | None:
    digest = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "mdmint"})
        with urllib.request.urlopen(req, timeout=30) as resp:  # noqa: S310 (user-provided URL)
            data = resp.read()
            ctype = resp.headers.get_content_type()
    except Exception as exc:
        log.warning("image: could not download %s (%s); leaving as-is.", url, exc)
        return None

    ext = _EXT_BY_TYPE.get(ctype) or os.path.splitext(url.split("?")[0])[1] or ".img"
    out_path = work_dir / f"remote-{digest}{ext}"
    out_path.write_bytes(data)
    return str(out_path)


def _resolve_one(src: str, base_dir: Path, work_dir: Path) -> str:
    low = src.lower()
    if low.startswith("data:"):
        return src
    if low.startswith(("http://", "https://")):
        local = _download(src, work_dir)
        return local if local else src
    if low.startswith("file://"):
        return urllib.request.url2pathname(src[7:])

    # Local path: absolute as-is, relative against the markdown file's directory.
    path = Path(src)
    if not path.is_absolute():
        path = base_dir / path
    abspath = os.path.abspath(path)
    if not os.path.exists(abspath):
        log.warning("image: file not found: %s (from src %r)", abspath, src)
    return abspath


def resolve_images(html_body: str, base_dir: Path, work_dir: Path) -> str:
    """Rewrite every ``<img src>`` in *html_body* to an absolute local path."""

    def repl(m: re.Match) -> str:
        prefix, quote, src = m.group(1), m.group(2), m.group(3)
        resolved = _resolve_one(src.strip(), base_dir, work_dir)
        # Keep the original quote style; the path itself contains no quote chars.
        return f"{prefix}{quote}{resolved}{quote}"

    return _IMG_SRC.sub(repl, html_body)
