"""mdmint — export Markdown files to DOCX and/or PDF.

Pure-Python, cross-platform, minimal dependencies. The two features that cannot
be rendered statically in pure Python — mermaid diagrams and LaTeX math — use an
optional renderer if present and fall back gracefully (with a warning) otherwise.

Public API
----------
    from mdmint import convert, MarkdownConverter

    convert(["doc.md"], formats=["pdf", "docx"], output="out/")
    convert_mkdocs("mkdocs.yml", formats=["docx"], template="brand.docx")
"""

from .core import MarkdownConverter, convert, convert_mkdocs

__version__ = "0.2.0rc1"
__all__ = ["MarkdownConverter", "convert", "convert_mkdocs", "__version__"]
