# CLAUDE.md — Document Rebranding to Ichita Style Tool

## Overview

Two scripts for producing Ichita-branded Word documents:

1. **`md_to_docx.py`** — Convert Markdown `.md` files to branded `.docx`
2. **`rebrand_skt_ichita.py`** — Rebrand an existing `.docx` (SKT pilot report) to Ichita style

## Font System

Both scripts use the same Thai/English split-run approach:

| Role | Font | Scale |
|---|---|---|
| English | Aeonik (auto-detected, fallback: Avenir Next) | 1.0x |
| Thai | Bai Jamjuree | 0.9x (e.g. Thai 9pt / English 10pt) |
| Code | Courier New | — |

**Why split-run?** Word's `szCs` (Complex Script size) doesn't work for Thai (Thai is not BiDi). Each text segment is split into separate `<w:r>` elements: Thai chars get Bai Jamjuree at scaled size, Latin chars get Aeonik at base size.

Key helpers: `_is_thai(c)`, `_split_thai_latin(text)`, `split_run_thai_latin(run_elem)` (in rebrander), `_add_split_run()` (in md_to_docx).

## Brand Colors

| Color | Hex | Usage |
|---|---|---|
| Primary Blue | #2978FF | Accent lines, H3-H4, links |
| Blue Grey 03 | #263338 | Body text, H1-H2, table headers |
| Blue Grey 02 | #788F9C | Captions, subtitles |
| Table Alt | #EFF2F3 | Alternating row shading |

## Usage

```bash
# Markdown to branded DOCX
python3 md_to_docx.py input.md              # → input.docx
python3 md_to_docx.py input.md output.docx  # custom output

# Rebrand existing DOCX
python3 rebrand_skt_ichita.py               # uses hardcoded SOURCE/OUTPUT paths
```

## Dependencies

```bash
pip3 install python-docx>=1.1.0
```
