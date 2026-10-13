"""Shared colour table for Markdown admonitions (``!!!`` / ``???`` callouts).

Both writers need to agree on what a "warning" looks like, but they style it
through completely different mechanisms: :mod:`mdmint.render` emits inline CSS that
the PDF backend honours, while :mod:`mdmint.backends` paints the DOCX cell with
low-level OOXML after the fact. Keeping the palette here means the two paths can
never drift, and adding a type is a one-line change.

Type names follow python-markdown / mkdocs-Material conventions; unknown tags
degrade to the neutral ``note`` styling rather than raising.
"""

from __future__ import annotations

# Canonical type -> (accent colour, light background tint), both ``#rrggbb``.
_CANONICAL: dict[str, tuple[str, str]] = {
    "note": ("#448aff", "#eaf1ff"),
    "abstract": ("#00b0ff", "#e5f6ff"),
    "info": ("#00b8d4", "#e3f8fc"),
    "tip": ("#00bfa5", "#e3f8f4"),
    "success": ("#00c853", "#e6f9ec"),
    "question": ("#64dd17", "#eef9e3"),
    "warning": ("#ff9100", "#fff3e0"),
    "failure": ("#ff5252", "#ffecec"),
    "danger": ("#ff1744", "#ffe9ed"),
    "bug": ("#f50057", "#ffe6ef"),
    "example": ("#7c4dff", "#f0ebff"),
    "quote": ("#9e9e9e", "#f5f5f5"),
}

# Extra spellings mkdocs/python-markdown accept, mapped to a canonical key.
_ALIASES: dict[str, str] = {
    "summary": "abstract",
    "tldr": "abstract",
    "hint": "tip",
    "important": "tip",
    "check": "success",
    "done": "success",
    "help": "question",
    "faq": "question",
    "caution": "warning",
    "attention": "warning",
    "fail": "failure",
    "missing": "failure",
    "error": "danger",
    "cite": "quote",
}

DEFAULT = "note"


def resolve(tag: str) -> str:
    """Map any admonition tag (or alias) to a canonical key, defaulting to note."""
    tag = (tag or "").strip().lower()
    if tag in _CANONICAL:
        return tag
    return _ALIASES.get(tag, DEFAULT)


def style_for(tag: str) -> tuple[str, str]:
    """Return ``(accent, background)`` hex colours for an admonition tag."""
    return _CANONICAL[resolve(tag)]


# Accent colour (upper-case, no ``#``) -> canonical type. The DOCX post-pass uses
# this to recognise an admonition table by its title colour, which is the only
# marker that survives htmldocx's class-stripping conversion.
ACCENTS: dict[str, str] = {
    accent.lstrip("#").upper(): key for key, (accent, _bg) in _CANONICAL.items()
}
