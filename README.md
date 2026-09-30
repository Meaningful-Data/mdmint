# mdup

Export Markdown files to **DOCX** and/or **PDF** — pure-Python, cross-platform,
minimal dependencies.

A single Markdown → HTML pass feeds both writers. The two features that can't be
rendered statically in pure Python (mermaid diagrams and LaTeX math) use an
**optional** renderer if one is available and fall back gracefully otherwise — so
the base install stays small and works everywhere.

## Install

```bash
pip install mdup
```

The core install is pure-Python (no native libraries):
`markdown-it-py`, `mdit-py-plugins`, `htmldocx`/`python-docx`, `xhtml2pdf`.

Optional extras:

```bash
pip install "mdup[math]"        # render LaTeX math (adds matplotlib)
pip install "mdup[weasyprint]"  # higher-fidelity PDF (needs native Pango/Cairo)
pip install "mdup[all]"         # both of the above
```

For **mermaid diagrams**, install the mermaid CLI separately (optional):
`npm install -g @mermaid-js/mermaid-cli` (provides `mmdc`).

## Usage

```bash
mdup notes.md                          # -> notes.pdf
mdup notes.md -f docx,pdf -o out/      # both formats into ./out/
mdup a.md b.md -f docx                 # -> a.docx and b.docx (one per input)
mdup a.md b.md --merge -f pdf -o book.pdf   # one merged PDF
mdup report.md --toc -f pdf            # insert a table of contents
mdup report.md -f pdf --pdf-backend weasyprint
```

| Option | Meaning |
| ------ | ------- |
| `-f, --format` | `pdf`, `docx`, or both (comma-separated). Default `pdf`. |
| `-o, --output` | Output directory, or a file path. Default: alongside each input. |
| `--merge` | With multiple inputs, produce one merged file per format. |
| `--toc` | Insert a table of contents (or honour a `[TOC]` marker). |
| `--pdf-backend` | `xhtml2pdf` (default) or `weasyprint`. |
| `-v, --verbose` | Verbose logging. |

| `--mkdocs FILE` | Build one document from an mkdocs.yml `nav` (instead of INPUTs). |
| `--section-style` | With `--mkdocs`: `part` (default) or `heading` — see below. |
| `--template FILE` | Word template (`.docx`/`.dotx`) for DOCX output. |
| `--meta KEY=VALUE` | Value for a template `{{KEY}}` placeholder (repeatable). |
| `--style ROLE=STYLE` | Use a specific template style for a role (repeatable). |

Also runnable as `python -m mdup ...`.

### From an mkdocs site

```bash
mdup --mkdocs mkdocs.yml -f docx,pdf -o out/     # -> out/<site_name>.docx/.pdf
```

The `nav` decides which pages go in and in what order (without a `nav`, the
`docs_dir` tree is used, like mkdocs does). Sections become headings:

- `--section-style part` (default): a **top-level section** is a *part* divider
  that doesn't use up a heading level, so its pages keep their `#` as chapters;
  nested sections become `#`, `##`… and demote their pages to match.
- `--section-style heading`: every section is a heading one level above its pages.

Pages under a section titled *Annex(es)*, *Appendix* or *Appendices* become
appendix chapters. Links between pages (`../other/page.md#anchor`) become
in-document links; links to pages left out of the nav are kept as plain text.
Headings accept mkdocs `attr_list` syntax (`# Title {#id .part}`).

Settings can live in the mkdocs.yml itself (command-line options win):

```yaml
extra:
  mdup:
    template: templates/brand.docx     # relative to mkdocs.yml
    section_style: part
    status: Final                      # any other key fills {{STATUS}} etc.
    confidentiality: Internal
    styles: {code: "My Code Style"}    # optional role overrides
```

### Branded DOCX with a Word template

```bash
mdup report.md -f docx --template brand.docx --meta status=Draft
```

With `--template`, DOCX output is written *into a copy of the template*: its
cover, contents page, headers/footers, page setup and styles are kept, and the
content is appended at the end (or where a paragraph reading `{{CONTENT}}` sits).
Text like `{{TITLE}}`, `{{AUTHOR}}`, `{{DATE}}` anywhere in the body, headers or
footers is filled from, in increasing priority: mkdocs.yml (`site_name` → title,
`site_description` → subtitle, `site_author` → author, `extra.mdup`), the Markdown
front matter, and `--meta`. `date` defaults to the current month; placeholders
left without a value are blanked with a warning.

Each element uses the template's styles, picking the first that exists:

| Role (`--style` key) | Styles tried |
| --- | --- |
| headings | `Heading 1`…`Heading N` (deeper levels reuse the last) |
| `part` / `appendix` | `Part Title` / `Appendix Heading` (else `Heading 1`) |
| `code` / `inline_code` | `Code Block`, `Source Code`, `HTML Preformatted` / `Inline Code`, `Verbatim Char`, `HTML Code` |
| `quote` | `Callout`, `Quote` (blockquotes) |
| `note` / `warning` | `Callout Note` / `Callout Warning` (admonitions, `> [!NOTE]`) |
| `table` | the template's own (non-stock) table style, else `Table Grid` |
| `table_heading` / `table_text` | `Table Heading` / `Table Text` |
| `figure` / `caption` | `Figure` / `Caption` ("Figure N." from the image alt text; `Table: …` before a table gives "Table N.") |
| lists | `List Bullet`, `List Number` (real multi-level numbering, restarted per list) |
| footnotes | `Footnote Text`/`Footnote Reference` (created if missing) — real Word footnotes |

How to build a template that styles well — required styles, numbering, placeholders, and a
specimen file to test it — is in [docs/TEMPLATES.md](docs/TEMPLATES.md).

Headings are bookmarked so in-document links work, and the template's TOC field
(if any) picks the headings up when Word updates fields on open. Without
`--template`, DOCX output is unchanged (htmldocx).

### Library API

```python
from mdup import convert

convert(["report.md"], formats=["pdf", "docx"], output="out/")
convert(["a.md", "b.md"], formats=["pdf"], output="book.pdf", merge=True)
```

## Supported Markdown

CommonMark plus: GFM tables, task lists, strikethrough, footnotes, auto table of
contents, fenced code, [admonitions](https://squidfunk.github.io/mkdocs-material/reference/admonitions/)
(`!!! note` / `??? warning` callouts), embedded images (relative paths resolved
against the source file; remote URLs downloaded), LaTeX math, and mermaid diagrams.

## Notes & limitations

- **PDF backend.** `xhtml2pdf` (default) is pure-Python and runs anywhere, with a
  pragmatic subset of CSS. `weasyprint` produces nicer output but needs native
  Pango/Cairo libraries, so it is an opt-in extra.
- **Math** uses matplotlib's *mathtext* engine (a large LaTeX-math subset, no system
  LaTeX required) rendered to images. Full LaTeX environments (e.g. `align`) are not
  supported. Without matplotlib, math is shown as literal source with a warning.
- **Mermaid** requires the `mmdc` CLI. Without it, diagrams are shown as styled code
  blocks with a warning. A snap-installed `mmdc` is supported (mdup places its
  scratch files where snap confinement can read them).
- **Admonitions** render as a colour-coded callout box (tinted background + accent
  bar + coloured title) in both PDF and DOCX. Type icons are intentionally omitted:
  the pure-Python PDF backend can't embed an emoji font, so colour and the title
  carry the type instead. The collapsible `???`/`???+` blocks render fully expanded
  (static documents have nothing to collapse into).

## License

Licensed under the [Apache License, Version 2.0](LICENSE).

Copyright 2026 Meaningful Data.
