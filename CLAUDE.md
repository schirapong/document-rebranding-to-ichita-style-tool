# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Tools for producing Ichita-branded Word documents:

1. **`rebrand_skt_ichita.py`** — Rebrand any existing DOCX to Ichita style (XML-level)
2. **`md_to_docx.py`** — Markdown → branded DOCX converter
3. **`install_fonts.py`** — Install brand fonts (Aeonik + Bai Jamjuree)

## Quick Start

```bash
# 1. Install dependencies
pip3 install -r requirements.txt

# 2. Install brand fonts (Aeonik + Bai Jamjuree)
python3 install_fonts.py           # install all fonts
python3 install_fonts.py --check   # check what's installed

# 3. Rebrand a DOCX
python3 rebrand_skt_ichita.py input.docx                # → input_ichita.docx
python3 rebrand_skt_ichita.py input.docx output.docx    # custom output
python3 rebrand_skt_ichita.py input.docx --logo my.png  # custom logo

# 4. Markdown → branded DOCX
python3 md_to_docx.py input.md              # → input.docx
python3 md_to_docx.py input.md output.docx  # custom output
```

## Architecture

### Font System

Brand fonts are bundled in the repo:
- **Aeonik** woff2 files in `Aeonik-Essentials-Web/` — converted to OTF by `install_fonts.py`
- **Bai Jamjuree** — downloaded from Google Fonts by `install_fonts.py` (OFL licensed)

`install_fonts.py` installs to the platform-appropriate font directory:
- macOS: `~/Library/Fonts/`
- Linux: `~/.local/share/fonts/`
- Windows: `%LOCALAPPDATA%/Microsoft/Windows/Fonts/`

### Split-Run Thai/English Font System

Word's `szCs` (Complex Script size) is ignored for Thai (Thai is not BiDi). The only way to control Thai font/size independently is separate `<w:r>` elements.

| Role | Font | Size |
|---|---|---|
| English | Aeonik (auto-detected, fallback: Avenir Next) | base size |
| Thai | Bai Jamjuree | base × 0.9 (e.g. 9pt when English 10pt) |
| Code | Courier New | base size |

### rebrand_skt_ichita.py — Document Rebrander (CLI)

**CLI usage**: `python3 rebrand_skt_ichita.py <input> [output] [--logo path]`

Deep-copies source DOCX body, then applies 20-step pipeline at XML level:
1. Copy body elements + image relationships
2. Remove old headers/footers
3. Cleanup (remove source logo, tagline, collapse empty paragraphs)
4. Restyle paragraphs — detects heading type by style ID sets or text pattern
5. Restyle tables — dark headers, alternating rows, brand borders
6. Reorder image captions above tables
7. Enforce spacing around tables
8. Redesign title page
9. Set margins, page break control
10. Add header/footer with ICHITA logo

**Heading detection**: Uses style ID sets (`KNOWN_TITLE_STYLES`, `KNOWN_SUBTITLE_STYLES`, `KNOWN_TOPIC_STYLES`) for common Thai templates, plus `detect_heading()` regex for text-based detection.

**Logo resolution**: Defaults to bundled `ichita brand ID/.../Wordmark/Ichita_Logo-05.png`. Override with `--logo`.

### md_to_docx.py — Markdown Converter

Line-by-line markdown parser → python-docx document builder.

## Spacing Rules

| Type | space_before | space_after |
|---|---|---|
| Section heading | 280 twips (14pt) | 120 twips (6pt) |
| Subsection | 200 twips (10pt) | 120 twips (6pt) |
| Caption | 120 twips (6pt) | 60 twips (3pt) |
| Body text | 0 | 120 twips (6pt) |

## Brand Colors

| Color | Hex | Usage |
|---|---|---|
| Primary Blue | #2978FF | Accent lines, H3-H4, links |
| Blue Grey 03 | #263338 | Body text, H1-H2, table headers |
| Blue Grey 02 | #788F9C | Captions, subtitles |
| Table Alt | #EFF2F3 | Alternating row shading |
| Table Border | #A0B0B8 | Subtle grid borders |

## Dependencies

```bash
pip3 install -r requirements.txt
# python-docx>=1.1.0  — DOCX generation/manipulation
# fonttools>=4.0.0     — woff2 → OTF conversion
# brotli>=1.0.0        — woff2 decompression
```
