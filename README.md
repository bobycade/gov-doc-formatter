# gov-doc-formatter / 政府公文格式化工具

Chinese government document formatter — converts markdown to `.docx` per **GB/T 9704-2012**.

政府公文格式化工具 — 将 Markdown 转换为符合 **GB/T 9704-2012** 标准的 `.docx` 文件。

## Features / 功能特性

- **标准合规排版**：A4 版面，正确页边距（3.7/3.5/2.8/2.6 cm），29pt 固定行距
- **多级标题体系**：一、/（一）/ 1. /（1）/ ①②③ 层次结构
- **智能模式（--auto）**：自动识别纯文本、`.docx`、`.wps`、`.doc` 或无格式 `.md` 的结构
- **三是标签加粗**：自动将"一是/二是/三是"等序数词及其后续内容（至首个句号）加粗
- **表格支持**：表头蓝底、黑体加粗，表身仿宋
- **中英混排**：中文文本中的英文/数字自动使用 Times New Roman

## Quick Start / 快速开始

```bash
# Standard: markdown → docx / 标准模式
python scripts/format_doc.py --input document.md

# Auto mode: .txt / .docx / .wps / .doc / unformatted .md → docx / 智能模式
python scripts/format_doc.py --input document.docx --auto
python scripts/format_doc.py --input document.wps --auto
```

## Requirements / 环境要求

- Python 3.8+
- `python-docx` (`pip install python-docx`)
- `pandoc`（auto 模式下处理 `.docx` 输入时需要）
- 处理 `.wps`/`.doc` 文件时需额外安装以下之一：**WPS Office** (Windows) 或 **LibreOffice**（跨平台）
- 系统字体：方正小标宋简体, 黑体, 楷体_GB2312, 仿宋_GB2312

## Usage / 使用方式

```
python scripts/format_doc.py --input source.md --output result.docx
python scripts/format_doc.py --input source.md --output result.docx --page-numbers
python scripts/format_doc.py --input notes.txt --auto
python scripts/format_doc.py --input messy.docx --auto
python scripts/format_doc.py --input document.wps --auto
python scripts/format_doc.py --input legacy.doc --auto
```

See [SKILL.md](SKILL.md) for full markdown syntax reference and formatting details.

完整 Markdown 语法参考及排版细节请参阅 [SKILL.md](SKILL.md)。

## License / 许可证

MIT
