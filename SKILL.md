---
name: gov-doc-formatter
description: "Formats markdown text into standard Chinese government document style (.docx) according to GB/T 9704-2012. Use when generating or reformatting official reports, speeches, or notices."
---

# Chinese Government Document Formatter

Converts markdown to .docx in standard Chinese official document (公文) format per GB/T 9704-2012.

## Quick Start

```bash
# Standard: markdown → docx
python scripts/format_doc.py --input document.md --output document.docx

# Auto mode: .txt / .docx / unformatted .md → docx
python scripts/format_doc.py --input document.docx --auto
python scripts/format_doc.py --input notes.txt --auto --output result.docx
```

If `--output` is omitted, the script auto-names the output. In auto mode, `_已排版` is appended.
Add `--page-numbers` to include page number footers.

### Auto Mode (`--auto`)

When the input is not a structured markdown file, use `--auto` to let the script detect heading hierarchy automatically:

| Input pattern | Detected as |
|--------------|-------------|
| First short line (no numbering) | `#` Main title |
| `一、二、三、…` at line start | `##` Level 1 chapter heading |
| `（一）（二）（三）…` at line start | `###` Level 2 section heading |
| `1. 2. 3. …` after a `###` section | `####` Level 3 sub-section heading |
| `一是/二是/第一/（1）/1.1` etc. | Body text (kept as-is) |

Long heading lines (> 45 chars with `。`) are auto-split: heading text stays as heading, trailing body text becomes a separate paragraph.

Supports .txt, .docx (extracted via pandoc), and unformatted .md files.

## Supported Markdown Features

| Markdown | Rendered as |
|----------|------------|
| `# Title` | Main title: 方正小标宋简体 22pt, centered |
| `> 版本：...` | Subtitle: 楷体\_GB2312 16pt, centered (supports `**bold**`) |
| `> **机构名**` | Bold subtitle: 楷体\_GB2312 16pt bold, centered |
| `## 一、...` | Chapter heading: 黑体 16pt bold, first-line indent 32pt |
| `### ...` | Section heading: 楷体\_GB2312 16pt bold, first-line indent 32pt. Auto-normalized to （一）（二）（三） |
| `#### ...` | Sub-section heading: 仿宋\_GB2312 16pt bold, first-line indent 32pt. Auto-normalized to 1. 2. 3. |
| Regular paragraph | Body: 仿宋\_GB2312 16pt, first-line indent 32pt |
| `**bold text**` | Bold span (works in all contexts: body, headings, notes, bullets, tables) |
| `"quoted"` | Auto-converted to Chinese quotation marks ""; ASCII single quotes `''` → `''` |
| `> quoted text` | Note: 楷体\_GB2312 14pt |
| `- list item` | Auto-numbered per GB/T 9704: `1. 2. 3.` under `##`/`###`, or `（1）（2）（3）`under `####`. Inside bold body labels, shifts one level deeper. |
| `+ list item` | Level 5 circled numbers: `①②③④⑤` (for feature callouts, grade breakdowns) |
| `1. list item` | Ordered list: 仿宋 16pt, first-line indent 32pt (passed through as-is) |
| `**numbered** body` | Bold body label: acts as structural divider. `1. 2. 3.` under `##`/`###`, `（1）（2）` under `####`. Subsequent `- ` items shift down one level. |
| `>> 发文机关` | Signature block: 仿宋 16pt, right-aligned (issuing authority + date) |
| `\| table \|` | Table: header 黑体 12pt bold (blue bg), body 仿宋 12pt |

English letters and numbers within Chinese text automatically use Times New Roman.

## Formatting Details

### Page Layout
- Paper: A4 (21.0 × 29.7 cm)
- Margins: top 3.7 cm, bottom 3.5 cm, left 2.8 cm, right 2.6 cm
- Line spacing: 29 pt (fixed)
- Body indent: 2 characters (32 pt at size 3)
- Justification: both sides

### Heading Hierarchy

| Markdown | GB/T 9704 Level | Font | Size | Style |
|----------|-----------------|------|------|-------|
| `#` | Main title | 方正小标宋简体 | 22pt | Centered (not bold) |
| `## 一、` | Level 1: 一、二、三、 | 黑体 | 16pt | Bold, 32pt indent |
| `### （一）` | Level 2: （一）（二）（三） | 楷体\_GB2312 | 16pt | Bold, 32pt indent |
| `#### 1.` | Level 3: 1. 2. 3. | 仿宋\_GB2312 | 16pt | Bold, 32pt indent |
| Body | — | 仿宋\_GB2312 | 16pt | Normal, 32pt indent |
| Note `>` | — | 楷体\_GB2312 | 14pt | Normal, 32pt indent |
| Signature `>>` | — | 仿宋\_GB2312 | 16pt | Right-aligned |

### Auto Numbering Normalization

The script automatically detects and corrects non-standard heading numbering:

- `###` headings: Leading numbers (1.1, 1., 一、, etc.) are stripped and replaced with sequential （一）（二）（三）... per chapter
- `####` headings: Leading numbers (4.2.1, 1., 一、, etc.) are stripped and replaced with sequential 1. 2. 3. ... per section
- `- ` list items: Context-aware numbering — under `##`/`###` headings use `1. 2. 3.` (Level 3); under `####` headings use （1）（2）（3）(Level 4). Numbering follows strict GB/T 9704 hierarchy (no level skipping).

This means you can write `### 1.1 编制目的` or `### 一、编制目的` in markdown — both will render as `（一）编制目的` in the output.

### GB/T 9704 Numbering Hierarchy

The standard defines a fixed 4-level sequence (never skip or reorder):

| Level | Format | Usage |
|-------|--------|-------|
| 1 | 一、二、三、 | `##` chapter headings |
| 2 | （一）（二）（三） | `###` section headings |
| 3 | 1. 2. 3. | `####` sub-section headings; `- ` list items under `##`/`###` |
| 4 | （1）（2）（3） | `- ` list items under `####` |
| 5 | ①②③ | Rarely used; not auto-generated |

Key rules:
- Levels must be used sequentially, never skipped
- Space between number and text: `1. 文字` (1 space after `.`)
- No punctuation after `（一）` or `（1）`

### Tables
- Header row: 黑体 12pt bold, centered, blue background (#D9E2F3)
- Data rows: 仿宋\_GB2312 12pt
- Content-heavy tables (cells > 20 chars): non-first columns use left alignment
- All tables: Word "Table Grid" style, centered on page

## Limitations

- Nested tables are not supported
- Images are not supported
- Table column widths are auto-sized by Word (not explicitly set)
- Font availability depends on system: 方正小标宋简体, 黑体, 楷体\_GB2312, 仿宋\_GB2312 must be installed
