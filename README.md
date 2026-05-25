# gov-doc-formatter

Chinese government document formatter — converts markdown to `.docx` per **GB/T 9704-2012**.

## Features

- **Standard-compliant formatting**: A4 layout, correct margins (3.7/3.5/2.8/2.6 cm), 29pt line spacing
- **Multi-level headings**: 一、/（一）/ 1. / （1）/ ①②③ hierarchy
- **Auto mode**: Detects structure from plain text, `.docx`, or unformatted `.md`
- **Body-label bolding**: `一是/二是/三是` labels auto-bolded through first sentence
- **Tables**: Header blue background, 黑体 bold header, 仿宋 body
- **Font mixing**: Times New Roman for English/numbers within Chinese text

## Quick Start

```bash
# Standard: markdown → docx
python scripts/format_doc.py --input document.md

# Auto mode: .txt / .docx / unformatted .md → docx
python scripts/format_doc.py --input document.docx --auto
```

## Requirements

- Python 3.8+
- `python-docx` (`pip install python-docx`)
- `pandoc` (for auto mode with `.docx` input)
- System fonts: 方正小标宋简体, 黑体, 楷体_GB2312, 仿宋_GB2312

## Usage

```
python scripts/format_doc.py --input source.md --output result.docx
python scripts/format_doc.py --input source.md --output result.docx --page-numbers
python scripts/format_doc.py --input notes.txt --auto
python scripts/format_doc.py --input messy.docx --auto
```

See [SKILL.md](SKILL.md) for full markdown syntax reference and formatting details.

## License

MIT
