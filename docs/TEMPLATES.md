# Authoring a Word template for mdmint

`mdmint --template brand.docx` writes DOCX output *into a copy of* `brand.docx`.
Everything visual — fonts, colours, spacing, cover, headers, footers, page
setup — comes from the template. mdmint only decides **which named style** each
piece of content gets. A good template is therefore mostly a well-built set of
styles.

This guide covers what mdmint reads from a template, which style names it looks
for, and how to check the result. The Meaningful Data template in
`templates/word-template/` follows every recommendation here and is a good model.

- [Quick checklist](#quick-checklist)
- [Anatomy of a template](#anatomy-of-a-template)
- [Placeholders](#placeholders)
- [Where the content goes](#where-the-content-goes)
- [Style reference](#style-reference)
- [Lists and numbering](#lists-and-numbering)
- [Headings, parts and appendices](#headings-parts-and-appendices)
- [Tables](#tables)
- [Callouts](#callouts)
- [Figures and captions](#figures-and-captions)
- [Footnotes](#footnotes)
- [Table of contents and fields](#table-of-contents-and-fields)
- [Overriding style choices](#overriding-style-choices)
- [Testing a template](#testing-a-template)

## Quick checklist

A template works if it opens in Word. It works *well* when it has:

- [ ] `Normal` set to your body font, size, colour and spacing.
- [ ] `Heading 1`–`Heading 4` styled, with outline levels 1–4 (Word's defaults).
- [ ] `List Bullet` and `List Number` **linked to multi-level numbering**
      (at least 3 levels). See [Lists and numbering](#lists-and-numbering).
- [ ] A custom **table style** (e.g. `Brand Table`) with a styled header row.
- [ ] `Code Block` (paragraph) and `Inline Code` (character) styles.
- [ ] `Callout`, `Callout Note` and `Callout Warning` for blockquotes and
      admonitions.
- [ ] `Caption` and `Figure`.
- [ ] Optionally: `Part Title`, `Appendix Heading`, `Table Heading`,
      `Table Text`, `List Continue`, `Footnote Text`, `Footnote Reference`,
      `Hyperlink`.
- [ ] A cover page using `{{TITLE}}`, `{{SUBTITLE}}`, `{{DATE}}`… placeholders.
- [ ] A contents page with a Word **TOC field**, if you want one.
- [ ] Headers and footers, including `PAGE` / `NUMPAGES` fields if wanted.
- [ ] Document **body otherwise empty**, or a paragraph reading `{{CONTENT}}`
      where content should go.

Every style is optional. A missing style falls back to a sensible one (see the
[style reference](#style-reference)), so even Word's blank `Normal.dotm` works.
It just looks plain.

## Anatomy of a template

```
┌──────────────────────┐  section 1 (no page numbers, own margins/borders)
│ logo                 │
│ {{TITLE}}            │  ← placeholders, filled from metadata
│ {{SUBTITLE}}         │
│ {{AUTHOR}} {{DATE}}  │
├──────────────────────┤  section break (next page)
│ Contents             │  section 2 (page numbering starts at 1)
│ [TOC field]          │
│                      │
│ {{CONTENT}}          │  ← optional: content goes here; else at the very end
└──────────────────────┘
  header: logo            footer: {{TITLE}} · {{DATE}} · page X / Y
```

- **Keep the cover and contents in the body.** mdmint never deletes what is in
  the template. It adds content after it.
- **Use section breaks** to give the cover different headers/footers and
  numbering from the rest. Content is added to the **last** section, so that
  section's header, footer, margins and page numbering apply to all generated
  pages.
- **Page size and margins** come from the last section. Images are scaled down
  to fit its text width (and 80% of its text height).
- `.docx` and `.dotx` are both accepted. Macros in `.docm`/`.dotm` are not run.

## Placeholders

Any text of the form `{{NAME}}` in the body, in a table (e.g. a cover grid), in
a text box, or in any header/footer is replaced with the metadata value `name`
(case-insensitive).

| Placeholder | Filled from (last wins) |
| --- | --- |
| `{{TITLE}}` | mkdocs `site_name` → front matter `title` → `--meta title=…`. Defaults to the document's first heading |
| `{{SUBTITLE}}` | mkdocs `site_description` → front matter → `--meta` |
| `{{AUTHOR}}` | mkdocs `site_author` → front matter → `--meta` |
| `{{DATE}}` | front matter → `--meta`. Defaults to the current month, e.g. "September 2026" |
| anything else, e.g. `{{STATUS}}`, `{{CONFIDENTIALITY}}`, `{{CLIENT}}` | mkdocs `extra: mdmint:` → front matter → `--meta` |

- You can invent any placeholder name. It just needs a value from one of the
  sources.
- Placeholders with no value are **blanked**, and mdmint prints a warning naming
  them.
- `title`, `author` and `subtitle` also set the document properties (File →
  Info), which Word can show via `TITLE`/`AUTHOR` fields.

**Type each placeholder in one go.** Word sometimes splits text into several
runs (after spell-check, autocorrect or partial formatting). mdmint still finds a
split placeholder, but the *whole paragraph* then takes the formatting of its
first run. If the text of a placeholder line has mixed formatting, retype the
placeholder, or format the whole paragraph with a paragraph style instead of
direct formatting.

## Where the content goes

- By default, content is appended at the **end of the body**.
- If a paragraph contains only `{{CONTENT}}` (or `{{BODY}}`), content is placed
  there instead and the marker paragraph is removed. Use this when something
  must follow the content, e.g. a closing page.
- If there is anything visible before the insertion point (cover, contents),
  the content **starts on a new page**. With an empty template body it starts on
  page 1.

## Style reference

mdmint looks for each role's styles **in order** and uses the first one the
template defines. Names are matched case-insensitively and are the names shown
in Word's Styles pane.

| Role | Styles tried, in order | Type | Used for |
| --- | --- | --- | --- |
| body | `Normal` | ¶ | Paragraphs |
| headings | `Heading 1` … `Heading N`; a missing level reuses the closest level above it | ¶ | `#` … `######` |
| `part` | `Part Title` → `Heading 1` | ¶ | Part dividers (top-level mkdocs sections, `# X {.part}`) |
| `appendix` | `Appendix Heading` → `Heading 1` | ¶ | Appendix chapters (`# X {.appendix}`, pages under "Annexes") |
| `list_bullet` | `List Bullet` (+ `List Bullet 2`, `3`… if single-level) | ¶ | `-` lists |
| `list_number` | `List Number` (+ `List Number 2`, `3`…) | ¶ | `1.` lists |
| `list_continue` | `List Continue` → `List Paragraph` | ¶ | 2nd+ paragraph inside a list item |
| `code` | `Code Block` → `Source Code` → `HTML Preformatted` → `Macro Text` | ¶ | Fenced code (one paragraph per block, lines joined by line breaks) |
| `inline_code` | `Inline Code` → `Verbatim Char` → `HTML Code` | char | `` `code` `` |
| `quote` | `Callout` → `Quote` → `Block Text` | ¶ | `> blockquote`, `!!! quote` |
| `note` | `Callout Note` → `Callout` → `Intense Quote` → `Quote` | ¶ | `!!! note/info/tip/…`, `> [!NOTE]`, `> [!TIP]` |
| `warning` | `Callout Warning` → `Callout` → `Intense Quote` → `Quote` | ¶ | `!!! warning/danger/bug/failure`, `> [!WARNING]`, `> [!CAUTION]` |
| `table` | first **custom** table style → `Table Grid` | table | Markdown tables |
| `table_heading` | `Table Heading` | ¶ | Header-row cells (else bold `Normal`) |
| `table_text` | `Table Text` | ¶ | Body cells |
| `figure` | `Figure` → `Captioned Figure` | ¶ | Paragraph holding a standalone image (else centred) |
| `caption` | `Caption` | ¶ | "Figure N." / "Table N." captions |
| `hyperlink` | `Hyperlink` | char | Links (else blue, underlined) |
| `footnote_text` | `Footnote Text` → `Reference` (else created) | ¶ | Footnote bodies |
| `footnote_ref` | `Footnote Reference` → `Reference Mark` (else created) | char | Footnote numbers, in text and in the note |
| `toc_heading` | `TOC Heading` | ¶ | Only if mdmint inserts a TOC itself (see below) |

Direct formatting that mdmint adds on top of styles is kept minimal: **bold**,
*italic*, ~~strike~~, super/subscript from the Markdown, left indents for list
continuation and code inside lists, 12 pt space before the paragraph that
follows a table, and 9 pt after the last item of a list.

**Base your styles on each other** (`Callout Note` based on `Callout`,
`Appendix Heading` on `Heading 1`), so a change to the parent flows through.

## Lists and numbering

This is the part templates most often get wrong. Bullets and numbers must come
from **Word numbering linked to the style**, not typed characters.

**Recommended: one multi-level list per style.**

1. Define a multi-level list (Home → Multilevel List → Define New) with at least
   3 levels: e.g. `•`, `–`, `·` for bullets; `1.`, `a.`, `i.` for numbers. Give
   each level its own indent (e.g. 0.25", 0.5", 0.75").
2. Link level 1 of that list to the `List Bullet` style (and a second list to
   `List Number`).

mdmint reads the numbering from the style. It uses level 2 and 3 of the *same*
list for nested items, and gives **every numbered list its own instance**, so
each list restarts at 1 (or at the Markdown start number, e.g. `3.`).

**Also supported: Word's stock layout**, where `List Bullet`, `List Bullet 2`,
`List Bullet 3` each have their own single-level numbering. mdmint then uses
`List Bullet 2` for the second level, and so on. With this layout, nested
numbered lists do not restart.

Continuation paragraphs (a second paragraph in the same item) use
`List Continue`, indented 18 pt per level. Make your list indents multiples of
18 pt (0.25") so the text lines up.

## Headings, parts and appendices

- **Outline levels** matter: the TOC and the Navigation pane are built from
  them. Keep `Heading 1` = level 1, etc. Give `Part Title` and
  `Appendix Heading` outline level 1 too, if they should appear in the TOC.
- **Page breaks:** set *Page break before* on `Heading 1` (and `Part Title`) if
  chapters should start on a new page. mdmint does not add breaks around
  headings itself.
- **Part Title** is used for top-level mkdocs nav sections
  (`--section-style part`) and for `# Title {.part}`. It is typically a divider
  page: page break before, large type, lots of space above.
- **Appendix Heading** is used for pages under an *Annex/Annexes/Appendix/
  Appendices* nav section, and for `# Title {.appendix}`. To number them
  "Appendix A", "Appendix B"…, link the style to a single-level list with format
  `A, B, C` and number text `Appendix %1 ·` (as the Meaningful Data template does).
- Every heading gets a bookmark, so links like `[see](#some-heading)` and
  cross-page mkdocs links become working internal hyperlinks.

## Tables

- Create a **custom table style** (Table Design → New Table Style). mdmint uses
  the first table style in the template that is not one of Word's stock ones.
  Delete stray custom table styles, or pin one with `--style table="Brand Table"`.
- mdmint turns on the style's **Header row**, **First column** and **Banded
  rows** options. Define *Header Row* conditional formatting (shading, bold,
  bottom border) and leave *First Column* unformatted unless you want it.
- The Markdown header row is marked "repeat as header row" on every page.
- Cell paragraphs use `Table Heading` / `Table Text`. Make these compact
  (spacing 0, slightly smaller than body text), because `Normal` spacing makes
  tables airy.
- Tables span the full text width. mdmint sizes columns from their content.
  Column alignment from Markdown (`|--:|`) is applied.

## Callouts

Blockquotes and admonitions become **consecutive paragraphs sharing one
style**, not boxes or tables. Build `Callout` / `Callout Note` /
`Callout Warning` from paragraph formatting:

- borders on all four sides, with the same colour and spacing on every
  paragraph,
- shading,
- left/right indent.

Word merges identical borders on consecutive paragraphs into **one box**. That
only works if every paragraph of the callout has the same borders and indents.
So use borders and shading, not a frame or text box. The admonition title is
the first paragraph, in bold.

## Figures and captions

- A standalone image (`![Alt text](diagram.png)` on its own line) goes into a
  `Figure` paragraph. Make `Figure` centred, with *Keep with next*, so the
  caption stays with the image.
- The alt text (or image title) becomes a caption: **Figure N.** *text*, in
  `Caption`. The number is a Word `SEQ Figure` field, so a *Table of Figures*
  works.
- A paragraph starting `Table: ` right before a table becomes a **Table N.**
  caption above it.
- Mermaid diagrams and math images get no caption.

## Footnotes

Markdown footnotes (`[^1]`) become real Word footnotes.

- With a `Footnote Text` style, the note reads `¹ text`.
- If your footnote style defines a **tab stop** (e.g. hanging indent 0.24" with
  a tab at 0.24"), mdmint puts a tab between the number and the text instead:
  `1⇥text`. This gives aligned, hanging notes.
- If neither `Footnote Text`/`Reference` nor `Footnote Reference`/
  `Reference Mark` exists, mdmint creates minimal ones (9 pt text, superscript
  reference).

## Table of contents and fields

- Put a **TOC field** on your contents page (References → Table of Contents, or
  a field `TOC \o "1-2" \h \z \u`). The `\o "1-2"` range controls the depth.
- mdmint never builds the contents list itself. Word fills it when fields are
  updated. mdmint switches on *update fields on open* whenever a TOC is present,
  so Word asks to update when the file opens (answer Yes). In LibreOffice, use
  Tools → Update → Update All.
- `PAGE`, `NUMPAGES`, `SECTIONPAGES` fields in headers/footers work as usual.
  Their cached values are stale until Word updates them.
- With `--toc`, if the template has **no** TOC field, mdmint inserts one (with a
  `TOC Heading` "Contents" title) where the content starts.

## Overriding style choices

If your template uses different names, map roles to styles without renaming
anything:

```bash
mdmint report.md -f docx --template brand.docx \
  --style code="Brand Code" --style note="Brand Info Box" --style table="Brand Grid"
```

or, for an mkdocs project, in `mkdocs.yml`:

```yaml
extra:
  mdmint:
    template: templates/brand.docx
    styles:
      code: Brand Code
      note: Brand Info Box
```

Role names are those in the [style reference](#style-reference) (`heading1`…
`heading6` for individual heading levels). A style name that doesn't exist is
reported and ignored.

## Testing a template

[`template-specimen.md`](template-specimen.md) uses every construct mdmint
styles. Render it with your template and page through the result:

```bash
mdmint docs/template-specimen.md -f docx --template brand.docx -o /tmp/specimen.docx
```

Check that:

1. The cover shows no leftover `{{…}}`, and mdmint printed no "no value for"
   warning you didn't expect.
2. Content starts on a new page after the contents page.
3. Nested bullets and numbers show three distinct levels, and the second
   numbered list restarts at 1.
4. The note, warning and quote callouts each form one box.
5. The table header is styled and repeats across pages; columns are readable.
6. Captions read "Figure 1." / "Table 1.".
7. Footnotes sit at the bottom of the page, with the numbering you expect.
8. After updating fields, the TOC lists parts, chapters and appendices.

To see which style every paragraph got:

```bash
python -c "import docx,sys; [print(f'{p.style.name:20} {p.text[:60]}') for p in docx.Document(sys.argv[1]).paragraphs]" /tmp/specimen.docx
```
