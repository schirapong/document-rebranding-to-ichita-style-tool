#!/usr/bin/env python3
"""
Convert Okara_Biorefinery_Technical_Process_v2.md to a professionally formatted .docx file.

Handles: headings (#, ##, ###, ####), tables (| delimited), code blocks (```),
bold (**text**), italic (*text*), bullet points (- item), numbered lists (1. item),
blockquotes (> text), horizontal rules (---), links [text](url), and regular paragraphs.
"""

import re
import os
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

INPUT_MD = "/Users/chirapongsakullachat/Documents/my-project/Okara_Biorefinery_Technical_Process_v2.md"
OUTPUT_DOCX = "/Users/chirapongsakullachat/Documents/my-project/Okara_Biorefinery_Technical_Process_v2.docx"


# -- Helpers -------------------------------------------------------------------

def set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_table_borders(table):
    """Apply borders to every cell in the table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)


def add_formatted_text(paragraph, text, base_font="Calibri", base_size=Pt(11), is_blockquote=False):
    """Parse inline markdown (bold, italic, bold+italic, links) and add runs to a paragraph."""
    # Pattern: ***bold+italic***, **bold**, *italic*, [link](url)
    pattern = re.compile(
        r'(\*\*\*(.+?)\*\*\*)'    # bold+italic
        r'|(\*\*(.+?)\*\*)'        # bold
        r'|(\*(.+?)\*)'            # italic
        r'|(\[([^\]]+)\]\(([^)]+)\))'  # link
    )

    last_end = 0
    for match in pattern.finditer(text):
        # Add text before this match
        before = text[last_end:match.start()]
        if before:
            run = paragraph.add_run(before)
            run.font.name = base_font
            run.font.size = base_size
            if is_blockquote:
                run.font.italic = True
                run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        if match.group(2):  # ***bold+italic***
            run = paragraph.add_run(match.group(2))
            run.font.bold = True
            run.font.italic = True
        elif match.group(4):  # **bold**
            run = paragraph.add_run(match.group(4))
            run.font.bold = True
            if is_blockquote:
                run.font.italic = True
                run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        elif match.group(6):  # *italic*
            run = paragraph.add_run(match.group(6))
            run.font.italic = True
        elif match.group(8):  # [link](url)
            link_text = match.group(9)
            run = paragraph.add_run(link_text)
            run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            run.font.underline = True

        run.font.name = base_font
        run.font.size = base_size
        last_end = match.end()

    # Remaining text
    remaining = text[last_end:]
    if remaining:
        run = paragraph.add_run(remaining)
        run.font.name = base_font
        run.font.size = base_size
        if is_blockquote:
            run.font.italic = True
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def add_cell_formatted_text(cell, text, bold_header=False, font_name="Calibri", font_size=Pt(10)):
    """Add formatted text to a table cell, preserving bold markdown."""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(2)

    if bold_header:
        # For header row, bold everything, strip markdown bold markers
        clean = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        run = paragraph.add_run(clean)
        run.font.bold = True
        run.font.name = font_name
        run.font.size = font_size
        run.font.color.rgb = RGBColor(0x1F, 0x2A, 0x44)
    else:
        # Parse bold markers within cell text
        pattern = re.compile(r'\*\*(.+?)\*\*')
        last_end = 0
        for match in pattern.finditer(text):
            before = text[last_end:match.start()]
            if before:
                run = paragraph.add_run(before)
                run.font.name = font_name
                run.font.size = font_size
            bold_text = match.group(1)
            run = paragraph.add_run(bold_text)
            run.font.bold = True
            run.font.name = font_name
            run.font.size = font_size
            last_end = match.end()
        remaining = text[last_end:]
        if remaining:
            run = paragraph.add_run(remaining)
            run.font.name = font_name
            run.font.size = font_size


def parse_table_line(line):
    """Parse a markdown table line into cell values."""
    cells = line.strip().strip('|').split('|')
    return [c.strip() for c in cells]


def is_separator_line(line):
    """Check if a line is a markdown table separator (|---|---|)."""
    stripped = line.strip().strip('|')
    parts = stripped.split('|')
    return all(re.match(r'^[\s\-:]+$', p) for p in parts)


def add_code_block(doc, code_lines):
    """Add a code block as a single paragraph with monospace font and grey background."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.right_indent = Inches(0.3)

    # Grey background shading on the paragraph
    pPr = p._p.get_or_add_pPr()
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F2F2" w:val="clear"/>')
    pPr.append(shading)

    code_text = '\n'.join(code_lines)
    run = p.add_run(code_text)
    run.font.name = 'Courier New'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)


def add_table_from_rows(doc, header_cells, data_rows):
    """Create a formatted table from parsed markdown table header and data rows."""
    num_cols = len(header_cells)

    # Normalize data rows to match header column count
    normalized_data = []
    for row in data_rows:
        while len(row) < num_cols:
            row.append("")
        normalized_data.append(row[:num_cols])

    table = doc.add_table(rows=1 + len(normalized_data), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    set_table_borders(table)

    # Header row
    hdr_row = table.rows[0]
    for j, cell_text in enumerate(header_cells):
        if j < num_cols:
            add_cell_formatted_text(hdr_row.cells[j], cell_text, bold_header=True)
            set_cell_shading(hdr_row.cells[j], "D6E4F0")  # Light blue

    # Data rows
    for row_idx, row_data in enumerate(normalized_data):
        row = table.rows[row_idx + 1]
        for j, cell_text in enumerate(row_data):
            if j < num_cols:
                add_cell_formatted_text(row.cells[j], cell_text)
                # Alternate row shading
                if row_idx % 2 == 1:
                    set_cell_shading(row.cells[j], "F5F5F5")

    # Small spacing after table
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(2)
    sp.paragraph_format.space_after = Pt(2)


# -- Main Conversion ----------------------------------------------------------

def convert_md_to_docx(input_path, output_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = Document()

    # Set default document font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.space_before = Pt(3)

    # Customize heading styles (1-4)
    heading_configs = [
        (1, 20, RGBColor(0x1A, 0x25, 0x5E)),
        (2, 16, RGBColor(0x1F, 0x4E, 0x79)),
        (3, 13, RGBColor(0x2E, 0x75, 0xB6)),
        (4, 11, RGBColor(0x2E, 0x75, 0xB6)),
    ]
    for level, size, color in heading_configs:
        hs = doc.styles[f'Heading {level}']
        hs.font.name = 'Calibri'
        hs.font.size = Pt(size)
        hs.font.color.rgb = color
        hs.font.bold = True
        if level == 1:
            hs.paragraph_format.space_before = Pt(18)
            hs.paragraph_format.space_after = Pt(8)
        elif level == 2:
            hs.paragraph_format.space_before = Pt(14)
            hs.paragraph_format.space_after = Pt(6)
        elif level == 3:
            hs.paragraph_format.space_before = Pt(10)
            hs.paragraph_format.space_after = Pt(6)
        else:
            hs.paragraph_format.space_before = Pt(8)
            hs.paragraph_format.space_after = Pt(4)

    # Set List Bullet style
    if 'List Bullet' in doc.styles:
        lb = doc.styles['List Bullet']
        lb.font.name = 'Calibri'
        lb.font.size = Pt(11)

    # Set document margins
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    i = 0
    first_h1 = True

    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip('\n')

        # -- Horizontal rule (---) --
        if re.match(r'^---+\s*$', stripped):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                '  <w:bottom w:val="single" w:sz="6" w:space="1" w:color="CCCCCC"/>'
                '</w:pBdr>'
            )
            pPr.append(pBdr)
            i += 1
            continue

        # -- Code block (```) --
        if stripped.startswith('```'):
            i += 1
            code_lines = []
            while i < len(lines) and not lines[i].rstrip('\n').startswith('```'):
                code_lines.append(lines[i].rstrip('\n'))
                i += 1
            if i < len(lines):
                i += 1  # skip closing ```
            add_code_block(doc, code_lines)
            continue

        # -- Table (| ... |) --
        if stripped.startswith('|') and '|' in stripped[1:]:
            table_lines = []
            while i < len(lines) and lines[i].rstrip('\n').strip().startswith('|'):
                table_lines.append(lines[i].rstrip('\n'))
                i += 1

            if len(table_lines) < 2:
                # Not a proper table, render as paragraphs
                for tl in table_lines:
                    p = doc.add_paragraph()
                    add_formatted_text(p, tl.strip())
                continue

            # Parse header row
            header_cells = parse_table_line(table_lines[0])

            # Find separator and data rows
            data_start = 1
            if len(table_lines) > 1 and is_separator_line(table_lines[1]):
                data_start = 2

            data_rows = []
            for tl in table_lines[data_start:]:
                if not is_separator_line(tl):
                    data_rows.append(parse_table_line(tl))

            add_table_from_rows(doc, header_cells, data_rows)
            continue

        # -- Headings (#, ##, ###, ####) --
        heading_match = re.match(r'^(#{1,4})\s+(.+)$', stripped)
        if heading_match:
            level = len(heading_match.group(1))
            heading_text = heading_match.group(2).strip()

            if level == 1 and first_h1:
                # Title: centered, larger font
                p = doc.add_heading('', level=1)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', heading_text)
                run = p.add_run(clean_text)
                run.font.name = 'Calibri'
                run.font.size = Pt(24)
                run.font.color.rgb = RGBColor(0x1A, 0x25, 0x5E)
                run.font.bold = True
                first_h1 = False
            else:
                p = doc.add_heading('', level=level)
                sizes = {1: Pt(20), 2: Pt(16), 3: Pt(13), 4: Pt(11)}
                colors = {
                    1: RGBColor(0x1A, 0x25, 0x5E),
                    2: RGBColor(0x1F, 0x4E, 0x79),
                    3: RGBColor(0x2E, 0x75, 0xB6),
                    4: RGBColor(0x2E, 0x75, 0xB6),
                }
                add_formatted_text(p, heading_text, base_font="Calibri", base_size=sizes[level])
                for run in p.runs:
                    run.font.color.rgb = colors[level]
                    run.font.bold = True

            i += 1
            continue

        # -- Blockquote (> text) --
        if stripped.startswith('>'):
            quote_text = stripped.lstrip('>').strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.right_indent = Inches(0.3)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(8)

            # Left border
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                '  <w:left w:val="single" w:sz="18" w:space="8" w:color="2E75B6"/>'
                '</w:pBdr>'
            )
            pPr.append(pBdr)

            # Light background
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EEF3FA" w:val="clear"/>')
            pPr.append(shading)

            add_formatted_text(p, quote_text, is_blockquote=True)
            i += 1
            continue

        # -- Numbered list (1. item) --
        numbered_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if numbered_match:
            num = numbered_match.group(1)
            item_text = numbered_match.group(2)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            run = p.add_run(f"{num}. ")
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run.font.bold = True

            add_formatted_text(p, item_text)
            i += 1
            continue

        # -- Bullet point (- item or * item, including nested) --
        bullet_match = re.match(r'^(\s*)([-*+])\s+(.+)$', stripped)
        if bullet_match:
            indent_spaces = len(bullet_match.group(1))
            bullet_text = bullet_match.group(3)
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)

            # Extra indent for nested bullets
            if indent_spaces >= 2:
                p.paragraph_format.left_indent = Inches(0.6 + (indent_spaces // 2) * 0.25)

            p.clear()
            add_formatted_text(p, bullet_text)
            i += 1
            continue

        # -- Empty line --
        if not stripped.strip():
            i += 1
            continue

        # -- Regular paragraph --
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(6)
        add_formatted_text(p, stripped)
        i += 1

    doc.save(output_path)

    # Verify output
    size = os.path.getsize(output_path)
    para_count = len(doc.paragraphs)
    table_count = len(doc.tables)
    print(f"Successfully saved: {output_path}")
    print(f"Output file size: {size:,} bytes ({size/1024:.1f} KB)")
    print(f"Document contains: {para_count} paragraphs, {table_count} tables")


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3:
        input_md = sys.argv[1]
        output_docx = sys.argv[2]
    else:
        input_md = INPUT_MD
        output_docx = OUTPUT_DOCX
    convert_md_to_docx(input_md, output_docx)
