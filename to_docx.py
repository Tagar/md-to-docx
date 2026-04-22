#!/usr/bin/env python3
"""Convert plain-text / light-markdown drafts into nicely-formatted Word .docx files.

Usage:
    to_docx.py INPUT.txt [OUTPUT.docx] [--title "Document title"]

Markdown features supported:
    Block-level:
    - `# H1`, `## H2`, `### H3`, `#### H4`                     explicit headings
    - `ALL CAPS LINE` or `1. ALL CAPS TITLE`                   implicit H1
    - Short `1. Title Case` (no end punctuation)               implicit H2
    - `N. Sentence with punctuation.`                          numbered list item
    - `- ` or `* ` (indent = sub-bullet)                       bullet
    - `> quoted text`                                          blockquote (multi-line ok)
    - ` ``` ` fenced code block                                monospace paragraph
    - Pipe tables (`| a | b |` + separator row)                docx table
    - Divider lines (`====` / `----`)                          dropped

    Inline (inside any paragraph/bullet/list/cell/heading):
    - `**text**` or `__text__`                                 bold
    - `*text*` or `_text_`                                     italic
    - `` `text` ``                                             monospace / inline code
    - `~~text~~`                                               strikethrough
    - `[label](url)`                                           hyperlink
    - `<html>` tags                                            stripped

    Label lines `Label: value`:                                bold label + value

Unit tests live in tests/test_to_docx.py.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


# ----- inline tokenization -----

# Order matters: `**` before `*`, `__` before `_`, tokens compete left-to-right.
INLINE_RE = re.compile(
    r"\*\*(?P<bold_ast>.+?)\*\*"                        # **bold**
    r"|__(?P<bold_und>.+?)__"                            # __bold__
    r"|\*(?P<italic_ast>[^*\n]+?)\*"                     # *italic*  (not **)
    r"|(?<![A-Za-z0-9_])_(?P<italic_und>[^_\n]+?)_(?![A-Za-z0-9_])"  # _italic_
    r"|`(?P<code>[^`\n]+?)`"                             # `inline code`
    r"|~~(?P<strike>.+?)~~"                              # ~~strike~~
    r"|\[(?P<link_text>[^\]]+)\]\((?P<link_url>[^)\s]+)\)"  # [text](url)
)

HTML_TAG_RE = re.compile(r"<!--.*?-->|<[^>]+>", re.DOTALL)


def strip_html(text: str) -> str:
    return HTML_TAG_RE.sub("", text)


def _add_hyperlink(paragraph, url: str, text: str):
    """Attach a clickable hyperlink run to `paragraph`."""
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rPr.append(underline)
    new_run.append(rPr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def _style_code_run(run) -> None:
    run.font.name = "Consolas"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0xC7, 0x25, 0x4E)  # dusty-rose code color


def add_inline_runs(paragraph, text: str) -> None:
    """Add `text` to `paragraph`, parsing markdown inline tokens into styled runs."""
    if not text:
        return
    text = strip_html(text)
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])

        if m.group("bold_ast") is not None:
            paragraph.add_run(m.group("bold_ast")).bold = True
        elif m.group("bold_und") is not None:
            paragraph.add_run(m.group("bold_und")).bold = True
        elif m.group("italic_ast") is not None:
            paragraph.add_run(m.group("italic_ast")).italic = True
        elif m.group("italic_und") is not None:
            paragraph.add_run(m.group("italic_und")).italic = True
        elif m.group("code") is not None:
            run = paragraph.add_run(m.group("code"))
            _style_code_run(run)
        elif m.group("strike") is not None:
            run = paragraph.add_run(m.group("strike"))
            run.font.strike = True
        elif m.group("link_text") is not None:
            _add_hyperlink(paragraph, m.group("link_url"), m.group("link_text"))
        pos = m.end()

    if pos < len(text):
        paragraph.add_run(text[pos:])


# ----- block-level classifiers -----

def is_divider(line: str) -> bool:
    s = line.strip()
    return len(s) >= 8 and set(s) <= {"=", "-"}


def is_markdown_heading(line: str):
    m = re.match(r"^(#{1,4})\s+(.+?)\s*$", line)
    if not m:
        return None
    return len(m.group(1)), m.group(2)


def is_all_caps_heading(line: str) -> bool:
    s = line.strip()
    if not s or len(s) < 3 or len(s) > 80:
        return False
    letters = [c for c in s if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)


def is_numbered_heading(line: str) -> bool:
    stripped = line.strip()
    return (
        bool(re.match(r"^\d+\.\s+[A-Z]", stripped))
        and len(stripped) < 100
        and not stripped.endswith(("?", ".", ",", ";", ":"))
    )


def is_numbered_list_item(line: str) -> bool:
    return bool(re.match(r"^\s*\d+\.\s+", line))


def is_bullet(line: str) -> bool:
    stripped = line.lstrip()
    return stripped.startswith("- ") or stripped.startswith("* ")


def is_kv_label(line: str) -> bool:
    return bool(re.match(r"^[A-Z][a-zA-Z ]{0,25}:\s", line.strip()))


def is_blockquote(line: str) -> bool:
    return line.lstrip().startswith(">")


def is_code_fence(line: str) -> bool:
    return line.strip().startswith("```")


def is_table_row(line: str) -> bool:
    return line.lstrip().startswith("|") and line.rstrip().endswith("|")


def is_table_separator(line: str) -> bool:
    stripped = line.strip()
    return bool(re.match(r"^\|?[\s\-:|]+\|?$", stripped)) and "-" in stripped and "|" in stripped


def parse_table_row(line: str) -> list[str]:
    # Strip leading/trailing pipes, split on inner pipes, strip each cell
    inner = line.strip()
    if inner.startswith("|"):
        inner = inner[1:]
    if inner.endswith("|"):
        inner = inner[:-1]
    return [c.strip() for c in inner.split("|")]


# ----- block builders -----

def _add_table(doc, rows: list[list[str]]) -> None:
    if not rows:
        return
    n_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=n_cols)
    try:
        table.style = "Light Grid Accent 1"
    except KeyError:
        pass
    for r_idx, row in enumerate(rows):
        for c_idx in range(n_cols):
            cell = table.cell(r_idx, c_idx)
            cell.paragraphs[0].clear()
            text = row[c_idx] if c_idx < len(row) else ""
            para = cell.paragraphs[0]
            add_inline_runs(para, text)
            # Header row: bold all runs
            if r_idx == 0:
                for run in para.runs:
                    run.bold = True


def _add_blockquote(doc, lines: list[str]) -> None:
    text = " ".join(l.strip() for l in lines)
    try:
        p = doc.add_paragraph(style="Intense Quote")
    except KeyError:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(24)
    add_inline_runs(p, text)


def _add_code_block(doc, code: str) -> None:
    for code_line in code.splitlines() or [""]:
        p = doc.add_paragraph()
        run = p.add_run(code_line if code_line else " ")
        _style_code_run(run)


# ----- main conversion loop -----

def convert(src: Path, dst: Path, title: str) -> None:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(title, level=0)

    lines = src.read_text().splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]

        # Divider or blank -> skip
        if is_divider(line) or not line.strip():
            i += 1
            continue

        # Fenced code block
        if is_code_fence(line):
            i += 1
            buf: list[str] = []
            while i < len(lines) and not is_code_fence(lines[i]):
                buf.append(lines[i])
                i += 1
            _add_code_block(doc, "\n".join(buf))
            if i < len(lines):
                i += 1  # skip closing fence
            continue

        # Pipe table: first row + separator row pattern
        if is_table_row(line) and i + 1 < len(lines) and is_table_separator(lines[i + 1]):
            rows = [parse_table_row(line)]
            i += 2  # skip header + separator
            while i < len(lines) and is_table_row(lines[i]):
                rows.append(parse_table_row(lines[i]))
                i += 1
            _add_table(doc, rows)
            continue

        # Blockquote (can span multiple > lines)
        if is_blockquote(line):
            quote_lines: list[str] = []
            while i < len(lines) and is_blockquote(lines[i]):
                quote_lines.append(lines[i].lstrip().lstrip(">").lstrip())
                i += 1
            _add_blockquote(doc, quote_lines)
            continue

        # Explicit markdown heading
        md = is_markdown_heading(line)
        if md:
            level, heading_text = md
            h = doc.add_heading("", level=level)
            add_inline_runs(h, heading_text)
            i += 1
            continue

        # Implicit ALL CAPS / all-caps-numbered top heading
        if is_all_caps_heading(line):
            doc.add_heading(line.strip().title(), level=1)
            i += 1
            continue

        # Short Title-Case numbered heading
        if is_numbered_heading(line):
            h = doc.add_heading("", level=2)
            add_inline_runs(h, line.strip())
            i += 1
            continue

        # Numbered list item
        if is_numbered_list_item(line):
            text = re.sub(r"^\s*\d+\.\s+", "", line)
            try:
                p = doc.add_paragraph(style="List Number")
            except KeyError:
                p = doc.add_paragraph()
            add_inline_runs(p, text)
            i += 1
            continue

        # Bullet
        if is_bullet(line):
            text = re.sub(r"^\s*[-*]\s+", "", line)
            style = "List Bullet 2" if line.startswith(("  - ", "    - ")) else "List Bullet"
            try:
                p = doc.add_paragraph(style=style)
            except KeyError:
                p = doc.add_paragraph()
                p.add_run("• ")
            add_inline_runs(p, text)
            i += 1
            continue

        # `Label: value`
        if is_kv_label(line):
            label, _, rest = line.strip().partition(":")
            p = doc.add_paragraph()
            p.add_run(label + ":").bold = True
            if rest.strip():
                add_inline_runs(p, " " + rest.strip())
            i += 1
            continue

        # Default paragraph
        p = doc.add_paragraph()
        add_inline_runs(p, line)
        i += 1

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)
    print(f"wrote {dst} ({dst.stat().st_size} bytes)")


def default_title(src: Path) -> str:
    for line in src.read_text().splitlines():
        if line.strip():
            return line.strip()
    return src.stem


def main() -> None:
    ap = argparse.ArgumentParser(description="Convert plain-text drafts to .docx")
    ap.add_argument("input", type=Path, help="input .txt or .md file")
    ap.add_argument("output", type=Path, nargs="?", default=None,
                    help="output .docx (default: <input_dir>/<stem>.docx)")
    ap.add_argument("--title", default=None,
                    help="document title (default: first non-empty line of input)")
    args = ap.parse_args()

    if not args.input.exists():
        sys.exit(f"error: {args.input} does not exist")

    out = args.output or args.input.with_suffix(".docx")
    title = args.title or default_title(args.input)
    convert(args.input, out, title)


if __name__ == "__main__":
    main()
