"""Command-line interface for mdmint."""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence

from . import __version__
from .core import VALID_FORMATS, MarkdownConverter
from .mkdocs import SECTION_STYLES


def _parse_formats(value: str) -> list[str]:
    fmts = [f.strip().lower() for f in value.split(",") if f.strip()]
    bad = [f for f in fmts if f not in VALID_FORMATS]
    if bad:
        raise argparse.ArgumentTypeError(
            f"invalid format(s): {', '.join(bad)} (choose from {', '.join(VALID_FORMATS)})"
        )
    if not fmts:
        raise argparse.ArgumentTypeError("no formats given")
    return fmts


def _key_value(value: str) -> tuple[str, str]:
    key, sep, val = value.partition("=")
    if not sep or not key.strip():
        raise argparse.ArgumentTypeError(f"expected KEY=VALUE, got {value!r}")
    return key.strip(), val.strip()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mdmint",
        description="Export Markdown file(s) to DOCX and/or PDF.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  mdmint notes.md                         # -> notes.pdf\n"
            "  mdmint notes.md -f docx,pdf -o out/     # both formats into out/\n"
            "  mdmint a.md b.md --merge -o book.pdf    # one merged PDF\n"
            "  mdmint a.md b.md -f docx                # a.docx and b.docx\n"
            "  mdmint --mkdocs mkdocs.yml -f docx --template brand.docx -o out/\n"
            "                                          # one branded DOCX from the nav\n"
        ),
    )
    parser.add_argument("inputs", nargs="*", metavar="INPUT", help="Markdown file(s)")
    parser.add_argument(
        "-f",
        "--format",
        dest="formats",
        type=_parse_formats,
        default=["pdf"],
        metavar="FMT[,FMT]",
        help="output format(s): pdf, docx, or both (comma-separated). Default: pdf",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="output directory, or a file path. Default: alongside each input.",
    )
    parser.add_argument(
        "--merge",
        action="store_true",
        help="with multiple inputs, produce a single merged file per format",
    )
    parser.add_argument(
        "--toc",
        action="store_true",
        help="insert a table of contents (or honour a [TOC] marker)",
    )
    parser.add_argument(
        "--mkdocs",
        metavar="MKDOCS_YML",
        default=None,
        help="build one document from an mkdocs config's nav (instead of INPUTs)",
    )
    parser.add_argument(
        "--section-style",
        choices=SECTION_STYLES,
        default=None,
        help="with --mkdocs: top-level nav sections become part dividers (part, "
        "default) or every section is a heading above its pages (heading)",
    )
    parser.add_argument(
        "--template",
        metavar="DOCX",
        default=None,
        help="Word template (.docx/.dotx) for DOCX output: its cover, headers, "
        "footers and styles are used",
    )
    parser.add_argument(
        "--meta",
        metavar="KEY=VALUE",
        type=_key_value,
        action="append",
        default=[],
        help="value for a template {{KEY}} placeholder (repeatable); overrides "
        "front matter / mkdocs.yml",
    )
    parser.add_argument(
        "--style",
        metavar="ROLE=STYLE",
        type=_key_value,
        action="append",
        default=[],
        help="use template style STYLE for ROLE, e.g. code='My Code' (repeatable)",
    )
    parser.add_argument(
        "--pdf-backend",
        choices=("xhtml2pdf", "weasyprint"),
        default="xhtml2pdf",
        help="PDF engine. xhtml2pdf (default, pure-Python) or weasyprint (extra).",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="verbose logging")
    parser.add_argument("--version", action="version", version=f"mdmint {__version__}")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.mkdocs and args.inputs:
        parser.error("give either INPUT files or --mkdocs, not both")
    if not args.mkdocs and not args.inputs:
        parser.error("no INPUT files given (or use --mkdocs MKDOCS_YML)")
    if args.section_style and not args.mkdocs:
        parser.error("--section-style only applies with --mkdocs")
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
    )
    try:
        converter = MarkdownConverter(
            args.formats,
            pdf_backend=args.pdf_backend,
            merge=args.merge,
            toc=args.toc,
            template=args.template,
            meta=dict(args.meta),
            styles=dict(args.style),
            section_style=args.section_style,
        )
        if args.mkdocs:
            written = converter.convert_mkdocs(args.mkdocs, output=args.output)
        else:
            written = converter.convert(args.inputs, output=args.output)
    except (ValueError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except ImportError as exc:
        print(f"error: missing optional dependency — {exc}", file=sys.stderr)
        return 3

    for path in written:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
