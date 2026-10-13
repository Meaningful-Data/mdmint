"""Build one document from an mkdocs project: its ``nav`` decides what goes in and
in which order, and its section nesting decides the heading structure.

Each page is rendered through the normal pipeline, then stitched together here:

* **Sections** become headings. With ``section_style="part"`` (default) a
  top-level section is a *part* divider (``<h1 class="part">``) that does not use
  up a heading level, so its pages keep their own ``#`` as chapters; nested
  sections become ``h1``, ``h2``… and demote their pages accordingly. With
  ``section_style="heading"`` every section is a plain heading one level above
  its pages.
* **Appendices** — pages under a section titled *Annex(es)* / *Appendix* /
  *Appendices* get their chapter heading marked ``class="appendix"``.
* **Links** between pages (``../other/index.md#anchor``) are rewritten to
  in-document anchors; links to pages left out of the nav are unwrapped to text.
* **Ids** are namespaced per page (``p3-…``) so headings and footnotes from
  different pages never collide.
"""

from __future__ import annotations

import html
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import unquote

from . import render

log = logging.getLogger("mdmint")

SECTION_STYLES = ("part", "heading")
_APPENDIX_TITLE = re.compile(r"^\s*(annex(es)?|appendix|appendices)\b", re.IGNORECASE)
_ANCHOR = re.compile(r'<a\b([^>]*?)\bhref="([^"]*)"([^>]*)>(.*?)</a>', re.DOTALL)
_URL_SCHEME = re.compile(r"^[a-z][a-z0-9+.-]*:", re.IGNORECASE)
# Keys of ``extra.mdmint`` that configure mdmint rather than fill template placeholders.
_OPTION_KEYS = ("template", "section_style", "styles")


@dataclass
class NavEntry:
    kind: str  # "section" or "page"
    title: str | None
    depth: int  # number of enclosing sections
    path: Path | None = None  # pages only (absolute)
    appendix: bool = False  # inside an annex/appendix section


@dataclass
class Project:
    config_path: Path
    docs_dir: Path
    site_name: str
    entries: list
    meta: dict = field(default_factory=dict)
    options: dict = field(default_factory=dict)


# ------------------------------------------------------------------ config


def _yaml_loader():
    """A SafeLoader that tolerates mkdocs' custom tags instead of failing on them.

    mkdocs.yml routinely contains ``!!python/name:…`` (pymdownx fences) and
    ``!ENV`` / ``!relative`` tags; we never need their values, only the nav.
    """
    import yaml

    class Loader(yaml.SafeLoader):
        pass

    def plain(loader, _suffix, node):
        if isinstance(node, yaml.ScalarNode):
            return loader.construct_scalar(node)
        if isinstance(node, yaml.SequenceNode):
            return loader.construct_sequence(node)
        return loader.construct_mapping(node)

    Loader.add_multi_constructor("tag:yaml.org,2002:python/", lambda *a: None)
    Loader.add_multi_constructor("!", plain)
    return Loader


def load(config_path) -> Project:
    """Read an mkdocs.yml and resolve its nav into an ordered list of entries."""
    import yaml

    config_path = Path(config_path).resolve()
    if not config_path.is_file():
        raise FileNotFoundError(f"mkdocs config not found: {config_path}")
    try:
        text = config_path.read_text(encoding="utf-8")
        cfg = yaml.load(text, Loader=_yaml_loader()) or {}  # noqa: S506 (a SafeLoader)
    except yaml.YAMLError as exc:
        raise ValueError(f"{config_path}: invalid YAML — {exc}") from exc
    if not isinstance(cfg, dict):
        raise ValueError(f"{config_path}: expected a mapping at the top level")

    docs_dir = (config_path.parent / (cfg.get("docs_dir") or "docs")).resolve()
    if not docs_dir.is_dir():
        raise FileNotFoundError(f"docs_dir not found: {docs_dir}")

    entries: list[NavEntry] = []
    if cfg.get("nav"):
        _walk_nav(cfg["nav"], docs_dir, 0, False, entries)
    else:
        log.info("mkdocs: no nav in %s; using the docs_dir file tree", config_path.name)
        _walk_tree(docs_dir, 0, entries)
    if not any(e.kind == "page" for e in entries):
        raise ValueError(f"{config_path}: nav contains no Markdown pages")

    site_name = str(cfg.get("site_name") or config_path.parent.name)
    extra = (cfg.get("extra") or {}).get("mdmint") or {}
    if not isinstance(extra, dict):
        extra = {}
    meta = {"title": site_name}
    if cfg.get("site_description"):
        meta["subtitle"] = str(cfg["site_description"])
    if cfg.get("site_author"):
        meta["author"] = str(cfg["site_author"])
    meta.update(
        {
            str(k).lower(): str(v)
            for k, v in extra.items()
            if k not in _OPTION_KEYS and v is not None
        }
    )
    options = {k: extra[k] for k in _OPTION_KEYS if k in extra}
    if "template" in options:
        options["template"] = (config_path.parent / str(options["template"])).resolve()
    return Project(config_path, docs_dir, site_name, entries, meta, options)


def _walk_nav(items, docs_dir: Path, depth: int, appendix: bool, out: list) -> None:
    if isinstance(items, (str, dict)):
        items = [items]
    for item in items or []:
        if isinstance(item, str):
            _add_page(None, item, docs_dir, depth, appendix, out)
        elif isinstance(item, dict):
            for title, value in item.items():
                title = None if title is None else str(title)
                if isinstance(value, list):
                    in_appendix = appendix or bool(title and _APPENDIX_TITLE.match(title))
                    section = NavEntry("section", title, depth, appendix=in_appendix)
                    start = len(out)
                    out.append(section)
                    _walk_nav(value, docs_dir, depth + 1, in_appendix, out)
                    if not any(e.kind == "page" for e in out[start + 1 :]):
                        del out[start:]  # drop sections with nothing to show
                elif isinstance(value, str):
                    _add_page(title, value, docs_dir, depth, appendix, out)


def _add_page(title, ref: str, docs_dir: Path, depth, appendix, out) -> None:
    if _URL_SCHEME.match(ref) or ref.startswith("//"):
        log.info("mkdocs: skipping external nav link %s", ref)
        return
    path = (docs_dir / ref).resolve()
    if path.is_dir():
        path = next((path / n for n in ("index.md", "README.md") if (path / n).is_file()), path)
    if not path.is_file():
        log.warning("mkdocs: nav entry not found, skipping: %s", ref)
        return
    if path.suffix.lower() not in (".md", ".markdown"):
        log.info("mkdocs: skipping non-Markdown nav entry %s", ref)
        return
    out.append(NavEntry("page", title, depth, path, appendix))


def _walk_tree(directory: Path, depth: int, out: list) -> None:
    """Implicit mkdocs nav: index/README first, then files and dirs by name."""

    def key(p: Path):
        return (p.stem.lower() not in ("index", "readme"), p.name.lower())

    for child in sorted(directory.iterdir(), key=key):
        if child.name.startswith("."):
            continue
        if child.is_dir():
            title = child.name.replace("-", " ").replace("_", " ").strip().capitalize()
            start = len(out)
            out.append(NavEntry("section", title, depth))
            _walk_tree(child, depth + 1, out)
            if not any(e.kind == "page" for e in out[start + 1 :]):
                del out[start:]
        elif child.suffix.lower() in (".md", ".markdown"):
            out.append(NavEntry("page", None, depth, child.resolve()))


# ---------------------------------------------------------------- assembly


def _levels(entry: NavEntry, section_style: str) -> int:
    """Section: its heading level (0 = part). Page: how far to demote it."""
    if section_style == "heading":
        return entry.depth + 1 if entry.kind == "section" else entry.depth
    if entry.kind == "section":
        return entry.depth  # 0 → part divider, 1 → h1, …
    return max(0, entry.depth - 1)


def assemble(project: Project, render_page, *, section_style: str = "part") -> str:
    """Render every nav page with ``render_page(path) -> html`` and join them.

    Returns the combined HTML fragment (no TOC; the caller adds one if wanted).
    """
    if section_style not in SECTION_STYLES:
        raise ValueError(f"Unknown section_style {section_style!r}; use one of {SECTION_STYLES}")

    # Pass 1: render pages, namespace their ids, fix up heading levels.
    prefixes: dict[Path, str] = {}
    first_ids: dict[Path, str] = {}
    blocks: list[tuple[NavEntry, str]] = []
    for idx, entry in enumerate(project.entries):
        if entry.kind == "section":
            level = _levels(entry, section_style)
            cls = ' class="part"' if level == 0 else ""
            anchor = f"s{idx}-{render.slugify(entry.title or 'section')}"
            label = html.escape(entry.title or "")
            blocks.append(
                (entry, f'<h{max(level, 1)} id="{anchor}"{cls}>{label}</h{max(level, 1)}>\n')
            )
            continue

        prefix = f"p{idx}-"
        prefixes.setdefault(entry.path, prefix)
        body = render_page(entry.path)
        if not re.search(r"<h1\b", body):
            # No H1 of its own: use the nav title (or file name) as the page title.
            title = entry.title or entry.path.stem.replace("-", " ").replace("_", " ").capitalize()
            body = f'<h1 id="{render.slugify(title)}">{html.escape(title)}</h1>\n' + body
        body = render.namespace_ids(body, prefix)
        offset = _levels(entry, section_style)
        if entry.appendix and offset == 0:
            body = _mark_first_h1(body, "appendix")
        body = render.shift_headings(body, offset)
        heads = render.headings_from_html(body)
        if heads and heads[0][2]:
            first_ids.setdefault(entry.path, heads[0][2])
        blocks.append((entry, body))

    # Pass 2: rewrite links between pages now that every target id is known.
    parts = []
    for entry, body in blocks:
        if entry.kind == "page":
            body = _rewrite_links(body, entry.path, prefixes, first_ids)
        parts.append(body)
    return "\n".join(parts)


def _mark_first_h1(body: str, cls: str) -> str:
    def repl(m):
        attrs = m.group(1)
        if 'class="' in attrs:
            attrs = attrs.replace('class="', f'class="{cls} ', 1)
        else:
            attrs += f' class="{cls}"'
        return f"<h1{attrs}>"

    return re.sub(r"<h1\b([^>]*)>", repl, body, count=1)


def _rewrite_links(body: str, page: Path, prefixes: dict, first_ids: dict) -> str:
    def repl(m: re.Match) -> str:
        before, href, after, inner = m.groups()
        if not href or href.startswith("#") or _URL_SCHEME.match(href) or href.startswith("//"):
            return m.group(0)
        target, _, frag = html.unescape(href).partition("#")
        target_path = (page.parent / unquote(target)).resolve() if target else page
        if target_path.is_dir():
            target_path = target_path / "index.md"
        if target_path.suffix.lower() not in (".md", ".markdown"):
            return m.group(0)  # a relative link to some other file; leave it
        if target_path not in prefixes:
            log.info(
                "mkdocs: %s links to %s, which is not in the nav; keeping text only",
                page.name,
                target,
            )
            return inner
        anchor = f"{prefixes[target_path]}{frag}" if frag else first_ids.get(target_path)
        if not anchor:
            return inner
        return f'<a{before}href="#{html.escape(anchor)}"{after}>{inner}</a>'

    return _ANCHOR.sub(repl, body)
