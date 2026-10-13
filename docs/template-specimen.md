---
title: Template Specimen
subtitle: Every construct mdmint styles, for checking a Word template
author: mdmint
status: Draft
confidentiality: Internal
---

# Part one {.part}

# Chapter heading (Heading 1)

Body paragraph in `Normal`, with **bold**, *italic*, ~~strike~~, `inline code`,
H<sub>2</sub>O and E = mc<sup>2</sup> (raw HTML sub/sup), an
[external link](https://example.com), and an internal link to
[the tables section](#tables).[^first]

## Section heading (Heading 2)

### Subsection heading (Heading 3)

#### Minor heading (Heading 4)

##### Deeper heading (reuses the closest level)

## Lists

- First bullet
- Second bullet
    - Nested bullet (level 2)
        - Nested bullet (level 3)
- Bullet with a continuation paragraph.

    This second paragraph belongs to the same item (`List Continue`).

1. First step
2. Second step
    1. Sub-step (level 2)
        1. Sub-sub-step (level 3)
3. Third step

A second numbered list must restart at 1:

1. Restarted item
2. Another item

- [x] Done task
- [ ] Open task

## Code

```python
def greet(name: str) -> str:
    return f"Hello, {name}!"   # one Code Block paragraph
```

## Callouts

> A plain blockquote uses `Callout`. It can span
> several lines and paragraphs.
>
> Second paragraph of the same callout.

!!! note "A note admonition"
    Uses `Callout Note`. The title line is bold.

!!! warning
    Uses `Callout Warning`.

> [!TIP]
> GitHub-style callouts are supported too.

## Tables

Table: Column alignment and a numbered caption

| Item | Description | Amount |
|:-----|:------------|-------:|
| Alpha | A short description | 1 |
| Beta | A longer description that wraps onto more than one line in a narrow column | 1,234 |
| Gamma | **Bold** and `code` inside a cell | 56 |

## Figures

![Standalone image with a numbered caption](../tests/fixtures/logo.png)

The text after a figure continues normally.[^second]

# Appendix chapter {.appendix}

Appendix chapters use `Appendix Heading` (numbered "Appendix A", "Appendix B"…
if the style is linked to a numbering list).

[^first]: The first footnote, with a [link](https://example.com).
[^second]: The second footnote.
