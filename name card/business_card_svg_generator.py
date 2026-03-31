#!/usr/bin/env python3
"""
Generate print-ready SVG business cards for ICHITA.
Output: SVG files at exact print dimensions with 3mm bleed and crop marks.
Open in Adobe Illustrator → Save As .ai for print production.
"""

# ICHITA wordmark SVG path data (dark fill)
WORDMARK_PATHS = """<rect x="177.56" y="102.75" width="24.31" height="97.25" fill="{fill}"/>
<polygon points="526.05 135.16 420.7 135.16 420.7 102.75 396.38 102.75 396.38 200 420.7 200 420.7 159.48 526.05 159.48 526.05 200 550.37 200 550.37 102.75 526.05 102.75 526.05 135.16" fill="{fill}"/>
<rect x="566.58" y="102.75" width="24.31" height="97.25" fill="{fill}"/>
<polygon points="902.44 200 846.29 102.75 820.81 102.75 818.4 102.75 792.92 102.75 736.77 200 764.66 200 819.63 102.75 874.54 200 902.44 200" fill="{fill}"/>
<path d="M216.85,150.47v1.8c0,26.36,21.37,47.73,47.73,47.73h115.6s0-20.71,0-20.71H245.17c-2.21,0-4-1.79-4-4v-47.9c0-2.21,1.79-4,4-4h135.01v-20.65h-115.6c-26.36,0-47.73,21.37-47.73,47.73Z" fill="{fill}"/>
<polygon points="607.1 102.75 607.1 123.46 671.93 123.46 671.93 200 696.25 200 696.25 123.46 761.08 123.46 761.08 102.75 607.1 102.75" fill="{fill}"/>
<polygon points="882.66 105.46 883.3 105.46 886.58 118.65 891.07 118.65 894.42 105.46 895.06 105.46 895.06 118.65 897.91 118.65 897.91 102.83 892.43 102.83 889.15 116.01 888.58 116.01 885.23 102.83 879.74 102.83 879.74 118.65 882.66 118.65 882.66 105.46" fill="{fill}"/>
<polygon points="869.41 118.65 872.33 118.65 872.33 105.46 877.6 105.46 877.6 102.83 864.06 102.83 864.06 105.46 869.41 105.46 869.41 118.65" fill="{fill}"/>"""

# Arrow X symbol path data (rotated 90°)
ARROW_X_PATHS = """<g transform="rotate(90, 540, 540)">
<path d="M391.7,290h-74L462,540L317.7,790h74l113.2-196c19.3-33.4,19.3-74.5,0-107.9L391.7,290L391.7,290z" fill="{fill}"/>
<path d="M688.3,790h74L618,540l144.3-250h-74L575.1,486c-19.3,33.4-19.3,74.5,0,107.9L688.3,790L688.3,790z" fill="{fill}"/>
</g>"""

# Brand colours
BLUE = "#2978FF"
BLUE_GREY_01 = "#CFD9DB"
BLUE_GREY_02 = "#788F9C"
BLUE_GREY_03 = "#263338"
WHITE = "#FFFFFF"

BLEED = 3  # mm


def mm(v):
    """Convert mm to SVG units (1mm = 1 unit, viewBox in mm)."""
    return v


def crop_marks(w, h, bleed=3, mark_len=5, stroke_width=0.25):
    """Generate crop marks outside bleed area."""
    marks = []
    corners = [
        (bleed, bleed),
        (bleed + w, bleed),
        (bleed, bleed + h),
        (bleed + w, bleed + h),
    ]
    for cx, cy in corners:
        # Horizontal marks
        if cx == bleed:
            marks.append(f'<line x1="{cx - bleed - mark_len}" y1="{cy}" x2="{cx - bleed}" y2="{cy}" stroke="black" stroke-width="{stroke_width}"/>')
        else:
            marks.append(f'<line x1="{cx + bleed}" y1="{cy}" x2="{cx + bleed + mark_len}" y2="{cy}" stroke="black" stroke-width="{stroke_width}"/>')
        # Vertical marks
        if cy == bleed:
            marks.append(f'<line x1="{cx}" y1="{cy - bleed - mark_len}" x2="{cx}" y2="{cy - bleed}" stroke="black" stroke-width="{stroke_width}"/>')
        else:
            marks.append(f'<line x1="{cx}" y1="{cy + bleed}" x2="{cx}" y2="{cy + bleed + mark_len}" stroke="black" stroke-width="{stroke_width}"/>')
    return "\n".join(marks)


def svg_header(total_w, total_h):
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg"
     width="{total_w}mm" height="{total_h}mm"
     viewBox="0 0 {total_w} {total_h}">
<defs>
  <style>
    @font-face {{
      font-family: 'Aeonik';
      src: local('Aeonik'), local('Aeonik-Regular');
    }}
    text {{ font-family: 'Aeonik', 'Avenir Next', 'Calibri', sans-serif; }}
  </style>
</defs>
'''


def wordmark_group(x, y, height_mm, fill):
    """Place ICHITA wordmark. Original viewBox: 177 100 725 105."""
    vb_w, vb_h = 725, 105
    aspect = vb_w / vb_h
    w = height_mm * aspect
    h = height_mm
    paths = WORDMARK_PATHS.format(fill=fill)
    return f'<g transform="translate({x},{y}) scale({w/vb_w},{h/vb_h}) translate(-177,-100)">\n{paths}\n</g>'


def arrow_x_group(x, y, height_mm, fill):
    """Place Arrow X symbol. Original viewBox: 300 270 480 540."""
    vb_w, vb_h = 480, 540
    aspect = vb_w / vb_h
    w = height_mm * aspect
    h = height_mm
    paths = ARROW_X_PATHS.format(fill=fill)
    return f'<g transform="translate({x},{y}) scale({w/vb_w},{h/vb_h}) translate(-300,-270)">\n{paths}\n</g>'


def generate_front_landscape(name_en, name_th, title_en, title_th, email, phone):
    """Generate front side of landscape card (90x55mm)."""
    w, h = 90, 55
    tw, th = w + 2*BLEED, h + 2*BLEED
    bx, by = BLEED, BLEED  # card origin

    svg = svg_header(tw, th)

    # Background (extends into bleed)
    svg += f'<rect x="0" y="0" width="{tw}" height="{th}" fill="{WHITE}"/>\n'

    # Blue accent bar at bottom (1mm height)
    svg += f'<rect x="0" y="{by + h - 1}" width="{tw}" height="{1 + BLEED}" fill="{BLUE}"/>\n'

    # Wordmark top-left
    svg += wordmark_group(bx + 7, by + 7, 5, BLUE_GREY_03)

    # Arrow X symbol + descriptor top-right
    svg += arrow_x_group(bx + 63, by + 5.5, 4, BLUE)
    svg += f'<text x="{bx + 67}" y="{by + 9}" font-size="3" fill="{BLUE_GREY_02}" letter-spacing="0.3" font-weight="500">SEPARATION TECHNOLOGIES</text>\n'

    # Name
    svg += f'<text x="{bx + 7}" y="{by + 25}" font-size="5.5" font-weight="600" fill="{BLUE_GREY_03}" letter-spacing="0.1">{name_en}</text>\n'
    svg += f'<text x="{bx + 7}" y="{by + 30}" font-size="4.5" font-weight="500" fill="{BLUE_GREY_02}">{name_th}</text>\n'

    # Title
    svg += f'<text x="{bx + 7}" y="{by + 34.5}" font-size="3.2" fill="{BLUE_GREY_02}" letter-spacing="0.2" text-transform="uppercase">{title_en.upper()}</text>\n'
    svg += f'<text x="{bx + 7}" y="{by + 38}" font-size="3.2" fill="{BLUE_GREY_02}">{title_th}</text>\n'

    # Blue separator
    svg += f'<line x1="{bx + 7}" y1="{by + 40}" x2="{bx + 15}" y2="{by + 40}" stroke="{BLUE}" stroke-width="0.3" opacity="0.5"/>\n'

    # Contact info
    cy = by + 44
    svg += f'<text x="{bx + 7}" y="{cy}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">E</tspan>   {email}</text>\n'
    svg += f'<text x="{bx + 7}" y="{cy + 4}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">T</tspan>   {phone}</text>\n'
    svg += f'<text x="{bx + 7}" y="{cy + 8}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">W</tspan>   ichita.co.th</text>\n'

    # Addresses (right-aligned)
    ax = bx + 83
    svg += f'<text x="{ax}" y="{cy}" font-size="2.8" fill="{BLUE_GREY_02}" text-anchor="end">399/75 Prachachuen Rd., Soi Pongpetchnivet,</text>\n'
    svg += f'<text x="{ax}" y="{cy + 3.2}" font-size="2.8" fill="{BLUE_GREY_02}" text-anchor="end">Jatujak, Bangkok 10900, Thailand</text>\n'
    svg += f'<line x1="{bx + 55}" y1="{cy + 5}" x2="{ax}" y2="{cy + 5}" stroke="{BLUE_GREY_01}" stroke-width="0.2" opacity="0.6"/>\n'
    svg += f'<text x="{ax}" y="{cy + 7.5}" font-size="3.2" fill="{BLUE_GREY_02}" text-anchor="end">399/75 ซอยพงษ์เพชรนิเวศน์ ถนนประชาชื่น</text>\n'
    svg += f'<text x="{ax}" y="{cy + 10.7}" font-size="3.2" fill="{BLUE_GREY_02}" text-anchor="end">เขตจตุจักร กรุงเทพฯ 10900</text>\n'

    # Crop marks
    svg += crop_marks(w, h)
    svg += "\n</svg>"
    return svg


def generate_back_landscape():
    """Generate back side of landscape card (90x55mm)."""
    w, h = 90, 55
    tw, th = w + 2*BLEED, h + 2*BLEED
    bx, by = BLEED, BLEED

    svg = svg_header(tw, th)

    # Dark background (extends into bleed)
    svg += f'<rect x="0" y="0" width="{tw}" height="{th}" fill="{BLUE_GREY_03}"/>\n'

    # Blue accent bar left
    svg += f'<rect x="0" y="0" width="{bx + 0.8}" height="{th}" fill="{BLUE}"/>\n'

    # Blue gradient glow
    svg += f'''<defs>
  <linearGradient id="glow" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0%" stop-color="{BLUE}" stop-opacity="0.08"/>
    <stop offset="100%" stop-color="{BLUE}" stop-opacity="0"/>
  </linearGradient>
</defs>\n'''
    svg += f'<rect x="0" y="0" width="{bx + 30}" height="{th}" fill="url(#glow)"/>\n'

    # Pattern bars (subtle)
    bar_y = by + h/2 - 15
    for i, bh in enumerate([10, 6, 3.5, 2, 1]):
        svg += f'<rect x="{bx + 3}" y="{bar_y}" width="{w - 6}" height="{bh}" fill="{WHITE}" opacity="0.06" rx="0.3"/>\n'
        bar_y += bh + 3.5

    # Arrow X symbol centered
    svg += arrow_x_group(bx + 35, by + 10, 20, WHITE)

    # Wordmark centered below
    svg += wordmark_group(bx + 28, by + 34, 5, WHITE)

    # Descriptor
    svg += f'<text x="{bx + w/2}" y="{by + 42}" font-size="2.5" fill="{BLUE_GREY_02}" letter-spacing="0.7" text-anchor="middle" font-weight="400">SEPARATION TECHNOLOGIES</text>\n'

    # Blue accent bar bottom
    svg += f'<rect x="0" y="{by + h - 1}" width="{tw}" height="{1 + BLEED}" fill="{BLUE}"/>\n'

    # Website
    svg += f'<text x="{bx + w - 7}" y="{by + h - 5}" font-size="2.8" fill="{BLUE}" font-weight="500" text-anchor="end" letter-spacing="0.15">ichita.co.th</text>\n'

    svg += crop_marks(w, h)
    svg += "\n</svg>"
    return svg


def generate_front_portrait(name_en, name_th, title_en, title_th, email, phone):
    """Generate front side of portrait card (55x90mm)."""
    w, h = 55, 90
    tw, th = w + 2*BLEED, h + 2*BLEED
    bx, by = BLEED, BLEED

    svg = svg_header(tw, th)

    # Background
    svg += f'<rect x="0" y="0" width="{tw}" height="{th}" fill="{WHITE}"/>\n'

    # Blue accent bar on left
    svg += f'<rect x="0" y="0" width="{bx + 1}" height="{th}" fill="{BLUE}"/>\n'

    # Wordmark
    svg += wordmark_group(bx + 6, by + 7, 5, BLUE_GREY_03)

    # Arrow X + descriptor below wordmark
    svg += arrow_x_group(bx + 6, by + 13, 3.2, BLUE)
    svg += f'<text x="{bx + 9.5}" y="{by + 15.5}" font-size="2.6" fill="{BLUE_GREY_02}" letter-spacing="0.25" font-weight="500">SEPARATION TECHNOLOGIES</text>\n'

    # Name (split across lines for long names)
    name_parts = name_en.split(" ", 1)
    svg += f'<text x="{bx + 6}" y="{by + 33}" font-size="6" font-weight="600" fill="{BLUE_GREY_03}" letter-spacing="0.1">{name_parts[0]}</text>\n'
    if len(name_parts) > 1:
        svg += f'<text x="{bx + 6}" y="{by + 39.5}" font-size="6" font-weight="600" fill="{BLUE_GREY_03}" letter-spacing="0.1">{name_parts[1]}</text>\n'
        name_th_y = by + 44
    else:
        name_th_y = by + 38
    svg += f'<text x="{bx + 6}" y="{name_th_y}" font-size="5" font-weight="500" fill="{BLUE_GREY_02}">{name_th}</text>\n'

    # Title
    title_y = name_th_y + 5.5
    svg += f'<text x="{bx + 6}" y="{title_y}" font-size="3.2" fill="{BLUE_GREY_02}" letter-spacing="0.2">{title_en.upper()}</text>\n'
    svg += f'<text x="{bx + 6}" y="{title_y + 3.8}" font-size="3.2" fill="{BLUE_GREY_02}">{title_th}</text>\n'

    # Blue separator
    sep_y = title_y + 7
    svg += f'<line x1="{bx + 6}" y1="{sep_y}" x2="{bx + 13}" y2="{sep_y}" stroke="{BLUE}" stroke-width="0.3" opacity="0.5"/>\n'

    # Contact info
    cy = sep_y + 4
    svg += f'<text x="{bx + 6}" y="{cy}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">E</tspan>   {email}</text>\n'
    svg += f'<text x="{bx + 6}" y="{cy + 4}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">T</tspan>   {phone}</text>\n'
    svg += f'<text x="{bx + 6}" y="{cy + 8}" font-size="3.2" fill="{BLUE_GREY_03}"><tspan fill="{BLUE}" font-weight="600" letter-spacing="0.2">W</tspan>   ichita.co.th</text>\n'

    # Addresses
    ay = cy + 13
    svg += f'<text x="{bx + 6}" y="{ay}" font-size="2.8" fill="{BLUE_GREY_02}">399/75 Prachachuen Rd., Soi Pongpetchnivet,</text>\n'
    svg += f'<text x="{bx + 6}" y="{ay + 3.2}" font-size="2.8" fill="{BLUE_GREY_02}">Jatujak, Bangkok 10900, Thailand</text>\n'
    svg += f'<line x1="{bx + 6}" y1="{ay + 5.2}" x2="{bx + 16}" y2="{ay + 5.2}" stroke="{BLUE_GREY_01}" stroke-width="0.2" opacity="0.6"/>\n'
    svg += f'<text x="{bx + 6}" y="{ay + 8}" font-size="3.2" fill="{BLUE_GREY_02}">399/75 ซอยพงษ์เพชรนิเวศน์ ถนนประชาชื่น</text>\n'
    svg += f'<text x="{bx + 6}" y="{ay + 11.5}" font-size="3.2" fill="{BLUE_GREY_02}">เขตจตุจักร กรุงเทพฯ 10900</text>\n'

    svg += crop_marks(w, h)
    svg += "\n</svg>"
    return svg


def generate_back_portrait():
    """Generate back side of portrait card (55x90mm)."""
    w, h = 55, 90
    tw, th = w + 2*BLEED, h + 2*BLEED
    bx, by = BLEED, BLEED

    svg = svg_header(tw, th)

    # Dark background
    svg += f'<rect x="0" y="0" width="{tw}" height="{th}" fill="{BLUE_GREY_03}"/>\n'

    # Blue accent bar left
    svg += f'<rect x="0" y="0" width="{bx + 1}" height="{th}" fill="{BLUE}"/>\n'

    # Gradient glow
    svg += f'''<defs>
  <linearGradient id="glowP" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0%" stop-color="{BLUE}" stop-opacity="0.08"/>
    <stop offset="100%" stop-color="{BLUE}" stop-opacity="0"/>
  </linearGradient>
</defs>\n'''
    svg += f'<rect x="0" y="0" width="{bx + 25}" height="{th}" fill="url(#glowP)"/>\n'

    # Pattern bars
    bar_y = by + h/2 - 18
    for i, bh in enumerate([13, 8, 4.5, 2.5, 1.2]):
        svg += f'<rect x="{bx + 3}" y="{bar_y}" width="{w - 6}" height="{bh}" fill="{WHITE}" opacity="0.06" rx="0.3"/>\n'
        bar_y += bh + 4.5

    # Arrow X symbol centered
    svg += arrow_x_group(bx + 17, by + 22, 25, WHITE)

    # Wordmark centered
    svg += wordmark_group(bx + 10, by + 54, 5.5, WHITE)

    # Descriptor
    svg += f'<text x="{bx + w/2}" y="{by + 63}" font-size="2.5" fill="{BLUE_GREY_02}" letter-spacing="0.7" text-anchor="middle" font-weight="400">SEPARATION TECHNOLOGIES</text>\n'

    # Blue accent bar bottom
    svg += f'<rect x="0" y="{by + h - 1}" width="{tw}" height="{1 + BLEED}" fill="{BLUE}"/>\n'

    # Website centered
    svg += f'<text x="{bx + w/2}" y="{by + h - 5}" font-size="2.8" fill="{BLUE}" font-weight="500" text-anchor="middle" letter-spacing="0.15">ichita.co.th</text>\n'

    svg += crop_marks(w, h)
    svg += "\n</svg>"
    return svg


def main():
    import os

    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "business_cards_print")
    os.makedirs(out_dir, exist_ok=True)

    people = [
        {
            "id": "chirapong",
            "name_en": "Chirapong Sakullachat",
            "name_th": "จิรพงษ์ สกุลชาติ",
            "title_en": "Chief Executive Officer",
            "title_th": "ประธานเจ้าหน้าที่บริหาร",
            "email": "chirapong@srithepgroup.com",
            "phone": "+66 81 841 7210",
        },
        {
            "id": "prakorn",
            "name_en": "Prakorn Makjamroen",
            "name_th": "ประกรณ์ เมฆจำเริญ",
            "title_en": "Business Advisor",
            "title_th": "ที่ปรึกษาธุรกิจ",
            "email": "prakorn@ichitathailand.com",
            "phone": "+66 81 789 2465",
        },
    ]

    for person in people:
        pid = person["id"]
        args = (person["name_en"], person["name_th"], person["title_en"],
                person["title_th"], person["email"], person["phone"])

        # Landscape
        with open(os.path.join(out_dir, f"{pid}_landscape_front.svg"), "w", encoding="utf-8") as f:
            f.write(generate_front_landscape(*args))
        with open(os.path.join(out_dir, f"{pid}_landscape_back.svg"), "w", encoding="utf-8") as f:
            f.write(generate_back_landscape())

        # Portrait
        with open(os.path.join(out_dir, f"{pid}_portrait_front.svg"), "w", encoding="utf-8") as f:
            f.write(generate_front_portrait(*args))
        with open(os.path.join(out_dir, f"{pid}_portrait_back.svg"), "w", encoding="utf-8") as f:
            f.write(generate_back_portrait())

        print(f"✓ {pid}: 4 SVGs (landscape front/back, portrait front/back)")

    print(f"\nAll files saved to: {out_dir}/")
    print("Open in Adobe Illustrator → File → Save As → .ai")
    print(f"\nPrint specs:")
    print(f"  Landscape: 90 × 55 mm + 3mm bleed")
    print(f"  Portrait:  55 × 90 mm + 3mm bleed")
    print(f"  Font: Aeonik (outline before sending to print)")
    print(f"  Colour: Convert to CMYK in Illustrator")


if __name__ == "__main__":
    main()
