"""Output writers: HTML → DOCX and HTML → PDF.

* DOCX uses ``htmldocx`` (built on ``python-docx``) — pure Python, no native deps.
* PDF uses ``xhtml2pdf`` by default (pure Python) or ``weasyprint`` if explicitly
  requested (higher fidelity, but needs native Pango/Cairo libraries).

Both PDF backends are imported lazily so a missing optional backend never affects
the default path.
"""

from __future__ import annotations

import os
import re
from importlib.resources import files as _res_files
from pathlib import Path
from typing import TYPE_CHECKING

from . import admonitions

if TYPE_CHECKING:
    from docx.document import Document as DocxDocument
    from docx.oxml.table import CT_TcPr

_IMG_SRC = re.compile(r'(<img\b[^>]*?\ssrc=)(["\'])(.*?)\2', re.IGNORECASE)


def load_css() -> str:
    """Return the bundled default stylesheet."""
    return _res_files("mdmint.resources").joinpath("default.css").read_text(encoding="utf-8")


def wrap_html(body: str, title: str = "", css: str | None = None) -> str:
    """Wrap an HTML body fragment in a complete, styled HTML document."""
    if css is None:
        css = load_css()
    safe_title = (title or "Document").replace("<", "&lt;").replace(">", "&gt;")
    return (
        "<!DOCTYPE html>\n<html><head>"
        '<meta charset="utf-8" />'
        f"<title>{safe_title}</title>"
        f"<style>{css}</style>"
        f"</head><body>{body}</body></html>"
    )


# --------------------------------------------------------------------------- DOCX


def _fit_images_to_page(document: DocxDocument) -> None:
    """Scale down any inline image larger than the page's text area.

    ``htmldocx`` adds pictures at their native size, and python-docx assumes 72 DPI
    for images without DPI metadata (mermaid/math PNGs), so wide diagrams overflow
    the margins. We shrink anything bigger than the printable area to fit, keeping
    aspect ratio and never upscaling — the DOCX counterpart of CSS ``max-width:100%``.
    """
    s = document.sections[0]
    max_w = (s.page_width or 0) - (s.left_margin or 0) - (s.right_margin or 0)
    max_h = (s.page_height or 0) - (s.top_margin or 0) - (s.bottom_margin or 0)
    if max_w <= 0 or max_h <= 0:
        return
    for shape in document.inline_shapes:
        w, h = shape.width, shape.height
        if not w or not h:
            continue
        scale = min(1.0, max_w / w, max_h / h)
        if scale < 1.0:
            shape.width = int(w * scale)
            shape.height = int(h * scale)


def _set_cell_shading(tcPr: CT_TcPr, fill: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _set_cell_borders(tcPr: CT_TcPr, accent: str) -> None:
    """A thick accent rule on the left, hairline grey on the other three sides."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    borders = OxmlElement("w:tcBorders")
    edges = {
        "left": ("24", accent),  # 24 eighths = 3pt accent bar
        "top": ("4", "DDDDDD"),
        "bottom": ("4", "DDDDDD"),
        "right": ("4", "DDDDDD"),
    }
    for edge, (size, color) in edges.items():
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), size)
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)
        borders.append(e)
    tcPr.append(borders)


def _style_admonitions(document: DocxDocument) -> None:
    """Paint admonition callouts in a freshly built DOCX.

    htmldocx renders our admonitions as bare single-cell tables and discards the
    CSS that colours them in PDF. We recognise them after the fact — a 1x1 table
    whose title run carries one of our accent colours (see :mod:`mdmint.admonitions`)
    — and add the matching cell shading and left accent bar. This is the DOCX twin
    of the inline CSS emitted by ``render._admonition_open``; anything that is not
    an admonition (real user tables) is left untouched.
    """
    for table in document.tables:
        if len(table.rows) != 1 or len(table.columns) != 1:
            continue
        cell = table.cell(0, 0)
        title = next((p for p in cell.paragraphs if p.text.strip()), None)
        run = title and next((r for r in title.runs if r.text.strip()), None)
        color = run and run.font.color
        if not color or color.rgb is None:
            continue
        accent = str(color.rgb).upper()
        if accent not in admonitions.ACCENTS:
            continue
        _accent, bg = admonitions.style_for(admonitions.ACCENTS[accent])
        tcPr = cell._tc.get_or_add_tcPr()
        _set_cell_shading(tcPr, bg.lstrip("#").upper())
        _set_cell_borders(tcPr, accent)


def _flatten_loose_list_items(body: str) -> str:
    """Unwrap the leading ``<p>`` inside each ``<li>`` for the DOCX path.

    markdown-it wraps list-item content in ``<p>`` whenever the list is *loose*
    (items separated by blank lines or spanning several lines). htmldocx turns
    such an ``<li>`` into two paragraphs — an empty ``List Number``/``List Bullet``
    paragraph carrying the marker, then a separate ``Normal`` paragraph with the
    text — so the number ends up on its own line. PDF is unaffected (CSS keeps the
    ``<p>`` inline), so this runs only here. Dropping the leading ``<p>`` wrapper
    puts the first line back beside its marker; any *further* paragraphs in a
    multi-block item are left as continuation paragraphs.
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(body, "html.parser")
    for li in soup.find_all("li"):
        first_p = li.find("p", recursive=False)
        if first_p is not None:
            first_p.unwrap()
    return str(soup)


def to_docx(
    body: str,
    out_path: str | os.PathLike[str],
    *,
    template: str | os.PathLike[str] | None = None,
    meta: dict[str, str] | None = None,
    styles: dict[str, str] | None = None,
) -> None:
    """Write *body* (an HTML fragment) to a .docx file.

    With a *template* (.docx/.dotx) the template-aware writer in
    :mod:`mdmint.docxtemplate` is used instead of htmldocx, so the output carries the
    template's cover, headers/footers and named styles (*meta* fills its
    ``{{PLACEHOLDERS}}``, *styles* overrides role → style-name choices).
    """
    if template is not None:
        from . import docxtemplate

        docxtemplate.to_docx(body, out_path, template, meta=meta, styles=styles)
        return

    from docx import Document
    from htmldocx import HtmlToDocx

    document = Document()
    parser = HtmlToDocx()
    parser.add_html_to_document(_flatten_loose_list_items(body), document)
    _fit_images_to_page(document)
    _style_admonitions(document)
    document.save(os.fspath(out_path))


# ---------------------------------------------------------------------------- PDF


def _xhtml2pdf_link_callback(uri: str, rel: str) -> str:
    """Resolve resource URIs for xhtml2pdf. Our image srcs are absolute paths."""
    if uri.startswith("file://"):
        import urllib.request

        return urllib.request.url2pathname(uri[7:])
    return uri


def _to_file_uris(body: str) -> str:
    """Convert absolute filesystem image paths to file:// URIs (for weasyprint)."""

    def repl(m: re.Match[str]) -> str:
        prefix, quote, src = m.group(1), m.group(2), m.group(3)
        if os.path.isabs(src) and os.path.exists(src):
            src = Path(src).as_uri()
        return f"{prefix}{quote}{src}{quote}"

    return _IMG_SRC.sub(repl, body)


def to_pdf(
    body: str,
    out_path: str | os.PathLike[str],
    *,
    backend: str = "xhtml2pdf",
    title: str = "",
    work_dir: Path | None = None,
) -> None:
    """Write *body* (an HTML fragment) to a .pdf file using the chosen backend."""
    if backend == "weasyprint":
        from weasyprint import HTML  # lazy: optional native-dependency backend

        html_doc = wrap_html(_to_file_uris(body), title=title)
        base = (work_dir or Path.cwd()).as_uri()
        HTML(string=html_doc, base_url=base).write_pdf(os.fspath(out_path))
        return

    if backend == "xhtml2pdf":
        from xhtml2pdf import pisa  # lazy

        html_doc = wrap_html(body, title=title)
        with open(os.fspath(out_path), "wb") as fh:
            status = pisa.CreatePDF(
                src=html_doc,
                dest=fh,
                encoding="utf-8",
                link_callback=_xhtml2pdf_link_callback,
            )
        if status.err:
            raise RuntimeError(f"xhtml2pdf failed to render {out_path}")
        return

    raise ValueError(f"Unknown PDF backend: {backend!r} (use 'xhtml2pdf' or 'weasyprint')")
