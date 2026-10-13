"""Markdown → HTML body, using markdown-it-py and a handful of plugins.

This is the single intermediate representation feeding both the DOCX and PDF
writers. Mermaid fences and LaTeX math are routed through the optional renderers
(:mod:`mdmint.mermaid`, :mod:`mdmint.mathrender`) via custom render rules; the work
directory they write images into is threaded through markdown-it's ``env``.
"""

from __future__ import annotations

import html
import re
import unicodedata
from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, Optional, cast

from markdown_it import MarkdownIt
from mdit_py_plugins.admon import admon_plugin
from mdit_py_plugins.anchors import anchors_plugin
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from mdit_py_plugins.front_matter import front_matter_plugin
from mdit_py_plugins.tasklists import tasklists_plugin

from . import admonitions, mathrender, mermaid

if TYPE_CHECKING:
    from markdown_it.renderer import RendererHTML
    from markdown_it.rules_core import StateCore
    from markdown_it.token import Token
    from markdown_it.utils import EnvType, OptionsDict

# A heading as (level, text, anchor id or None).
Heading = tuple[int, str, Optional[str]]

_TOC_MARKER = re.compile(r"<p>\[\[?TOC\]\]?</p>", re.IGNORECASE)

# Trailing attr_list on a heading, python-markdown / mkdocs style:
# ``# Title {#custom-id .part}`` or ``# Title {: .appendix }``.
_HEADING_ATTRS = re.compile(r"\s*\{:?\s*((?:[#.][\w-]+\s*)+)\}\s*$")


def slugify(title: str) -> str:
    """Heading id slug compatible with mkdocs' default ``toc`` slugify.

    Matching mkdocs matters because docs written for it link to headings by these
    ids (``page.md#2-the-maintenance-case``). Non-ASCII is folded away as mkdocs
    does; a heading with nothing ASCII left keeps its Unicode word characters so it
    still gets a usable id.
    """
    ascii_ = unicodedata.normalize("NFKD", title).encode("ascii", "ignore").decode("ascii")
    for value in (ascii_, title):
        value = re.sub(r"[^\w\s-]", "", value).strip().lower()
        value = re.sub(r"[-\s]+", "-", value)
        if value:
            return value
    return "section"


def _heading_attrs_strip(state: StateCore) -> None:
    """Core rule (before anchors): strip a trailing ``{#id .class}`` from headings.

    The attrs are parked in ``token.meta`` so the slug is computed from the clean
    title, then re-applied by :func:`_heading_attrs_apply` after the anchors rule
    (which would otherwise overwrite a custom id).
    """
    tokens = state.tokens
    for i, tok in enumerate(tokens):
        if tok.type != "heading_open" or i + 1 >= len(tokens):
            continue
        inline = tokens[i + 1]
        last = inline.children[-1] if inline.children else None
        if last is None or last.type != "text":
            continue
        m = _HEADING_ATTRS.search(last.content)
        if not m:
            continue
        last.content = last.content[: m.start()]
        inline.content = _HEADING_ATTRS.sub("", inline.content)
        parts = m.group(1).split()
        tok.meta = dict(tok.meta or {})
        tok.meta["attr_id"] = next((p[1:] for p in parts if p.startswith("#")), None)
        tok.meta["attr_classes"] = [p[1:] for p in parts if p.startswith(".")]


def _heading_attrs_apply(state: StateCore) -> None:
    for tok in state.tokens:
        if tok.type != "heading_open" or not tok.meta:
            continue
        if tok.meta.get("attr_id"):
            tok.attrSet("id", tok.meta["attr_id"])
        if tok.meta.get("attr_classes"):
            tok.attrSet("class", " ".join(tok.meta["attr_classes"]))


def _math_inline(tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType) -> str:
    return mathrender.render(tokens[idx].content, block=False, work_dir=env["work_dir"])


def _math_block(tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType) -> str:
    return mathrender.render(tokens[idx].content, block=True, work_dir=env["work_dir"])


# --- Admonitions (!!! / ??? callouts) -------------------------------------
#
# We render each admonition as a full-width single-cell table rather than the
# plugin's default <div>. A <div> with CSS looks right in PDF but htmldocx drops
# it entirely (no CSS, no box); a one-cell table is the one structure both
# writers turn into a real box. The dynamic per-type colours go inline so the PDF
# backend needs no generated stylesheet, and the accent-coloured title doubles as
# the marker the DOCX post-pass keys on (see backends._style_admonitions).


def _admonition_open(tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType) -> str:
    tok = tokens[idx]
    tag = (tok.meta or {}).get("tag", "")
    canon = admonitions.resolve(tag)
    accent, bg = admonitions.style_for(tag)
    env.setdefault("_admon_stack", []).append((accent, bg))
    classes = tok.attrGet("class") or f"admonition {canon}"
    return (
        f'<table class="{classes}"><tbody><tr>'
        f'<td class="admonition-cell" style="background-color:{bg};'
        f'border-left:4px solid {accent};">'
    )


def _admonition_close(tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType) -> str:
    stack = env.get("_admon_stack")
    if stack:
        stack.pop()
    return "</td></tr></tbody></table>\n"


def _admonition_title_open(
    tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType
) -> str:
    stack = env.get("_admon_stack") or [admonitions.style_for("note")]
    accent = stack[-1][0]
    # The colour must live on a <span> (not the <p>): htmldocx only copies run
    # colour from a span's inline style, and that coloured run is exactly what the
    # DOCX post-pass uses to recognise the admonition.
    return (
        f'<p class="admonition-title" style="margin:0 0 0.4em 0;">'
        f'<span style="color:{accent};"><strong>'
    )


def _admonition_title_close(
    tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType
) -> str:
    return "</strong></span></p>\n"


def _build_md() -> MarkdownIt:
    """Construct a configured MarkdownIt instance (stateless across renders)."""
    md = (
        MarkdownIt("commonmark", {"html": True, "linkify": True, "typographer": False})
        .enable(["table", "strikethrough"])
        .use(front_matter_plugin)  # strip YAML front-matter from output
        .use(footnote_plugin)
        .use(tasklists_plugin, enabled=True)
        .use(anchors_plugin, max_level=6, permalink=False, slug_func=slugify)
        .use(dollarmath_plugin, double_inline=True)
        .use(admon_plugin)  # !!! note / ??? collapsible callouts
    )
    md.core.ruler.before("anchor", "heading_attrs_strip", _heading_attrs_strip)
    md.core.ruler.after("anchor", "heading_attrs_apply", _heading_attrs_apply)

    # Renderer rules are called as rule(tokens, idx, options, env): plain functions work.
    rules: dict[str, Any] = cast("RendererHTML", md.renderer).rules

    # Route mermaid fenced blocks through the optional renderer; everything else
    # uses the default fence rule.
    default_fence = rules["fence"]

    def fence(tokens: Sequence[Token], idx: int, options: OptionsDict, env: EnvType) -> str:
        token = tokens[idx]
        info = token.info.strip()
        lang = info.split(maxsplit=1)[0] if info else ""
        if lang.lower() == "mermaid":
            return mermaid.render(token.content, env["work_dir"])
        return default_fence(tokens, idx, options, env)

    rules["fence"] = fence

    # Route math tokens through the optional renderer.
    rules["math_inline"] = _math_inline
    rules["math_inline_double"] = _math_block  # $$...$$ used inline
    rules["math_block"] = _math_block
    rules["math_block_label"] = _math_block

    # Route admonitions to a styled single-cell table (see helpers above).
    rules["admonition_open"] = _admonition_open
    rules["admonition_close"] = _admonition_close
    rules["admonition_title_open"] = _admonition_title_open
    rules["admonition_title_close"] = _admonition_title_close
    return md


_MD = _build_md()


def _inline_text(inline_token: Token | None) -> str:
    """Plain-text content of an inline token (for TOC labels)."""
    if not inline_token or not inline_token.children:
        return inline_token.content if inline_token else ""
    return "".join(c.content for c in inline_token.children if c.type in ("text", "code_inline"))


def _extract_headings(tokens: Sequence[Token]) -> list[Heading]:
    """Return (level, text, anchor_id) for every heading in the document."""
    headings: list[Heading] = []
    for i, tok in enumerate(tokens):
        if tok.type == "heading_open":
            level = int(tok.tag[1])
            anchor = tok.attrGet("id")
            text = _inline_text(tokens[i + 1]) if i + 1 < len(tokens) else ""
            headings.append((level, text, None if anchor is None else str(anchor)))
    return headings


def build_toc(headings: Sequence[Heading]) -> str:
    """Build a ``<nav class="toc">`` block from a list of (level, text, anchor)."""
    if not headings:
        return ""
    items = []
    for level, text, anchor in headings:
        label = html.escape(text)
        link = f'<a href="#{anchor}">{label}</a>' if anchor else label
        indent = (level - 1) * 1.2
        items.append(f'<li style="margin-left:{indent:.1f}em">{link}</li>')
    return f'<nav class="toc"><div class="toc-title">Contents</div><ul>{"".join(items)}</ul></nav>'


def has_toc_marker(body: str) -> bool:
    return bool(_TOC_MARKER.search(body))


def strip_toc_markers(body: str) -> str:
    """Remove any literal ``[TOC]`` markers from a rendered body."""
    return _TOC_MARKER.sub("", body)


def prefix_anchors(
    body: str, headings: Sequence[Heading], prefix: str
) -> tuple[str, list[Heading]]:
    """Namespace heading anchors so several documents can share one HTML page.

    Returns ``(body, headings)`` with each heading's ``id`` (and the matching TOC
    anchor) prefixed, so a consolidated TOC over merged files links unambiguously
    even when two files use the same heading text. Non-heading ids (e.g. footnotes)
    are left untouched.
    """
    mapping: dict[str, str] = {}
    new_headings: list[Heading] = []
    for level, text, anchor in headings:
        if anchor:
            mapping[anchor] = f"{prefix}{anchor}"
            new_headings.append((level, text, mapping[anchor]))
        else:
            new_headings.append((level, text, anchor))
    if mapping:

        def repl(m: re.Match[str]) -> str:
            old = m.group(2)
            return f"{m.group(1)}{mapping.get(old, old)}{m.group(3)}"

        body = re.sub(r'(\bid=")([^"]*)(")', repl, body)
    return body, new_headings


_HEADING_TAG = re.compile(r"<(/?)h([1-6])(?=[\s>])", re.IGNORECASE)
_HEADING_HTML = re.compile(r"<h([1-6])\b([^>]*)>(.*?)</h\1>", re.IGNORECASE | re.DOTALL)
_ID_ATTR = re.compile(r'\bid="([^"]*)"')


def shift_headings(body: str, offset: int) -> str:
    """Demote every heading in *body* by *offset* levels (capped at ``h6``)."""
    if not offset:
        return body
    return _HEADING_TAG.sub(lambda m: f"<{m.group(1)}h{min(6, int(m.group(2)) + offset)}", body)


def namespace_ids(body: str, prefix: str) -> str:
    """Prefix *every* ``id`` and in-page ``href="#..."`` in *body*.

    Unlike :func:`prefix_anchors` (headings only), this also covers footnotes, so
    several documents stitched into one keep their footnote links apart.
    """
    body = re.sub(
        r'(\bid=")([^"]*)(")', lambda m: f"{m.group(1)}{prefix}{m.group(2)}{m.group(3)}", body
    )
    return re.sub(
        r'(\bhref="#)([^"]+)(")', lambda m: f"{m.group(1)}{prefix}{m.group(2)}{m.group(3)}", body
    )


def headings_from_html(body: str) -> list[Heading]:
    """Return (level, text, anchor_id) for every heading in an HTML fragment."""
    out: list[Heading] = []
    for m in _HEADING_HTML.finditer(body):
        idm = _ID_ATTR.search(m.group(2))
        text = html.unescape(re.sub(r"<[^>]+>", "", m.group(3))).strip()
        out.append((int(m.group(1)), text, idm.group(1) if idm else None))
    return out


def parse(md_text: str, *, work_dir: Path) -> tuple[str, list[Heading]]:
    """Render Markdown to an HTML fragment and return ``(body, headings)``.

    ``headings`` is a list of (level, text, anchor) in document order. No table of
    contents is inserted — use :func:`apply_toc` (single file) or :func:`build_toc`
    (consolidated across merged files).
    """
    env: dict[str, Any] = {"work_dir": work_dir}
    tokens = _MD.parse(md_text, env)
    body = _MD.renderer.render(tokens, _MD.options, env)
    return body, _extract_headings(tokens)


def apply_toc(body: str, headings: Sequence[Heading], toc: bool) -> str:
    """Insert a single-document TOC: at a ``[TOC]`` marker if present, else at the
    top when *toc* is True. A stray marker with no TOC wanted is dropped."""
    toc_html = build_toc(headings)
    if toc_html and _TOC_MARKER.search(body):
        return _TOC_MARKER.sub(toc_html, body, count=1)
    if toc and toc_html:
        return toc_html + "\n" + body
    if _TOC_MARKER.search(body):
        return _TOC_MARKER.sub("", body)
    return body


def render_body(md_text: str, *, work_dir: Path, toc: bool = False) -> str:
    """Render Markdown source to an HTML fragment (no <html>/<body> wrapper).

    toc: if True, insert a table of contents at the top (unless a ``[TOC]`` marker
         is present, in which case it is inserted there regardless of this flag).
    """
    body, headings = parse(md_text, work_dir=work_dir)
    return apply_toc(body, headings, toc)
