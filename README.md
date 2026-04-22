# md-to-docx

A single-file Python script that converts plain-text / light-markdown drafts into nicely-formatted Microsoft Word `.docx` files.

Designed for the workflow of drafting in a plain text editor (or having an LLM write a draft) and then handing a clean, styled Word document to a stakeholder — without round-tripping through pandoc or a full markdown processor.

## Why

Most markdown-to-docx tools are heavyweight (pandoc) or target GitHub-flavored markdown where every edge case matters. This one is optimized for the common "write a scoping doc / briefing / memo" use case:

- Zero config — run the script, get a styled `.docx`
- Sensible defaults for headings, lists, tables, inline formatting
- Implicit heading detection (ALL CAPS, numbered titles) — works on drafts that aren't strictly markdown
- Pipe tables render as proper Word tables with the `Light Grid Accent 1` style
- Hyperlinks are clickable
- Inline code renders in monospace with a color accent

## Installation

```sh
pip install python-docx
# or with uv:
uv pip install python-docx
```

Clone the repo and run `to_docx.py` directly — no package install required.

## Usage

```sh
./to_docx.py draft.txt                    # -> draft.docx
./to_docx.py draft.md output.docx
./to_docx.py draft.txt --title "My Doc"   # override document title
```

By default, the output file is written next to the input with the `.docx` extension, and the document title is the first non-empty line of the input.

## Features

### Block-level

| Syntax | Result |
|---|---|
| `# H1`, `## H2`, `### H3`, `#### H4` | Explicit headings |
| `ALL CAPS LINE` | Implicit H1 |
| `1. Title Case` (short, no end punctuation) | Implicit H2 |
| `1. Sentence with punctuation.` | Numbered list item |
| `- item` or `* item` | Bullet (indent for sub-bullets) |
| `> quoted text` | Blockquote (multi-line ok) |
| ` ``` ` code fence | Monospace code block |
| `\| a \| b \|` + `\|---\|---\|` | Word table |
| `====` or `----` | Divider (dropped) |
| `Label: value` | Bold label + value |

### Inline

| Syntax | Result |
|---|---|
| `**text**` or `__text__` | **Bold** |
| `*text*` or `_text_` | *Italic* |
| `` `text` `` | `Inline code` |
| `~~text~~` | ~~Strikethrough~~ |
| `[label](url)` | Clickable hyperlink |
| `<html>` tags | Stripped |

Inline formatting works inside headings, list items, table cells, and blockquotes.

## Example

Given `draft.txt`:

```
My Scoping Doc

## Overview

This project will deliver **three** things:

- A data pipeline using `dbt`
- An ML model (*Prophet* baseline)
- A [dashboard](https://example.com)

## Timeline

| Phase | Weeks | Owner |
|---|---|---|
| Discovery | 1-2 | Lead |
| Build | 3-6 | Team |
| Handoff | 7-8 | PM |
```

Run:

```sh
./to_docx.py draft.txt
```

You get a `.docx` with a title page, styled headings, bulleted lists with inline formatting preserved, and a properly-formatted Word table.

See `examples/sample.txt` for a more complete input demonstrating every supported feature.

## Testing

```sh
pip install python-docx pytest
pytest tests/
```

The test suite covers the block-level classifiers, HTML stripping, and end-to-end conversion (generating real `.docx` files and asserting on their contents via `python-docx`).

## Design notes

- One file, no framework — `to_docx.py` is ~350 lines, readable end-to-end
- The inline tokenizer uses a single regex with named alternation groups; block-level classifiers are small predicate functions
- Pipe tables are detected by header-row-plus-separator pattern, not by a trailing state machine
- No dependencies beyond `python-docx`

## License

MIT — see `LICENSE`.
