#!/usr/bin/env python3
"""
Therefore Category Index Form Renderer

Renders Therefore XML configuration category index forms as PNG images.
Reads a Therefore Configuration XML export and generates visual representations
of each category's index form layout.

Usage:
    python render_categories.py <input_xml> <output_dir>

Example:
    python render_categories.py TheConfiguration.xml ./renders
"""

import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont
import textwrap
import os
import re
import sys

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
SCALE_X = 1.52
SCALE_Y = 1.70
PADDING_LEFT = 12
PADDING_TOP = 35

BG_COLOR = (240, 240, 240)
WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
BORDER_COLOR = (172, 172, 172)
HEADER_COLOR = (200, 200, 200)
TAB_ACTIVE_BG = (240, 240, 240)
TAB_INACTIVE_BG = (220, 220, 220)

# Known lookup field names that get a "..." button
LOOKUP_FIELDS = {'Supplier ID', 'Supplier Name', 'Supplier ABN'}

# ---------------------------------------------------------------------------
# XML Helpers
# ---------------------------------------------------------------------------
def get_text(elem, path):
    e = elem.find(path)
    return e.text if e is not None else None


def get_name(elem):
    name = elem.find('Name')
    if name is not None:
        tstr = name.find('TStr')
        if tstr is not None:
            t = tstr.find('T')
            if t is not None:
                s = t.find('S')
                return s.text if s is not None else None
    return None


def get_caption(field):
    cap = field.find('Caption')
    if cap is not None:
        tstr = cap.find('TStr')
        if tstr is not None:
            t = tstr.find('T')
            if t is not None:
                s = t.find('S')
                return s.text if s is not None else None
    return None


# ---------------------------------------------------------------------------
# Coordinate Scaling
# ---------------------------------------------------------------------------
def scale_x(v):
    return int(float(v) * SCALE_X) + PADDING_LEFT


def scale_y(v):
    return int(float(v) * SCALE_Y) + PADDING_TOP


def scale_w(v):
    return int(float(v) * SCALE_X)


def scale_h(v):
    return max(int(float(v) * SCALE_Y), 18)


# ---------------------------------------------------------------------------
# Field Filtering
# ---------------------------------------------------------------------------
def is_valid_field(field):
    """Skip fields that are hidden or have no meaningful size/position."""
    visible = get_text(field, 'Visible')
    if visible == '0':
        return False
    pos_x = get_text(field, 'PosX')
    pos_y = get_text(field, 'PosY')
    fwidth = get_text(field, 'Width')
    fheight = get_text(field, 'Height')
    if pos_x is None or pos_y is None:
        return False
    if float(pos_x) == 0 and float(pos_y) == 0:
        if fwidth is None or fheight is None or float(fwidth) == 0 or float(fheight) == 0:
            return False
    return True


def wrap_text(text, max_chars):
    if not text or len(text) <= max_chars:
        return [text] if text else []
    return textwrap.wrap(text, width=max_chars, break_long_words=False)


# ---------------------------------------------------------------------------
# Field Parsing
# ---------------------------------------------------------------------------
def get_table_columns(fields_list):
    """
    Collect zero-height / zero-position fields that define table columns.
    These are fields with PosX=0, PosY=0, Height=0 but Width>0.
    """
    cols = []
    for field in fields_list:
        px = float(get_text(field, 'PosX') or 0)
        py = float(get_text(field, 'PosY') or 0)
        h  = float(get_text(field, 'Height') or 0)
        w  = float(get_text(field, 'Width') or 0)
        tno = int(get_text(field, 'TypeNo') or 1)
        vis = get_text(field, 'Visible')
        if vis == '0':
            continue
        if px == 0 and py == 0 and h == 0 and w > 0 and tno not in (4, 13):
            cap = get_caption(field) or get_text(field, 'ColName') or ''
            cols.append({'caption': cap, 'width': w, 'type': tno})
    return cols


def get_fields_for_rendering(fields_list):
    """Parse a list of Field XML elements into labels and inputs."""
    # Collect table column definitions first
    table_columns = get_table_columns(fields_list)

    labels = []
    inputs = []
    for field in fields_list:
        if not is_valid_field(field):
            continue
        caption = get_caption(field) or get_text(field, 'ColName') or ''
        type_no = int(get_text(field, 'TypeNo') or 1)
        pos_x = get_text(field, 'PosX')
        pos_y = get_text(field, 'PosY')
        fwidth = get_text(field, 'Width')
        fheight = get_text(field, 'Height')
        x = scale_x(pos_x)
        y = scale_y(pos_y)
        w = scale_w(fwidth) if fwidth else 100
        h = scale_h(fheight) if fheight else 18
        raw_x = float(pos_x)
        raw_y = float(pos_y)

        if type_no == 4:
            labels.append({
                'caption': caption, 'x': x, 'y': y, 'w': w, 'h': h,
                'raw_x': raw_x, 'raw_y': raw_y,
            })
        elif type_no == 6:
            inputs.append({
                'caption': caption, 'x': x, 'y': y, 'w': w, 'h': h,
                'type': 'checkbox', 'raw_x': raw_x, 'raw_y': raw_y,
            })
        elif type_no == 10:
            inputs.append({
                'caption': caption, 'x': x, 'y': y, 'w': w, 'h': h,
                'type': 10, 'raw_x': raw_x, 'raw_y': raw_y,
                'columns': table_columns,
            })
        else:
            inputs.append({
                'caption': caption, 'x': x, 'y': y, 'w': w, 'h': h,
                'type': type_no, 'raw_x': raw_x, 'raw_y': raw_y,
            })
    return labels, inputs


# ---------------------------------------------------------------------------
# Drawing Helpers
# ---------------------------------------------------------------------------
def draw_table_field(draw, font, inp):
    """Render a TypeNo=10 table field with column header row."""
    x, y, w, h = inp['x'], inp['y'], inp['w'], inp['h']
    cols = inp.get('columns', [])

    # Outer table border
    draw.rectangle([x, y, x + w, y + h], fill=WHITE, outline=BORDER_COLOR)

    if not cols:
        draw.text((x + 4, y + 4), "(table)", fill=(150, 150, 150), font=font)
        return

    header_h = 16
    # Header background
    draw.rectangle([x, y, x + w, y + header_h], fill=HEADER_COLOR, outline=BORDER_COLOR)

    # Distribute columns proportionally by their XML Width values
    total_w = sum(c['width'] for c in cols)
    col_x = x + 1
    for i, col in enumerate(cols):
        col_w = int((col['width'] / total_w) * (w - 2))
        # Column header text
        bbox = draw.textbbox((0, 0), col['caption'], font=font)
        text_w = bbox[2] - bbox[0]
        # Centre-align short text, left-align if wider than column
        tx = col_x + max(2, (col_w - text_w) // 2)
        draw.text((tx, y + 3), col['caption'], fill=BLACK, font=font)
        # Column divider (not after last column)
        if i < len(cols) - 1:
            draw.line([col_x + col_w, y, col_x + col_w, y + header_h], fill=BORDER_COLOR)
        col_x += col_w

    # Row lines to suggest a grid (2 empty rows)
    row_h = max(16, (h - header_h) // 3)
    for r in range(1, 3):
        ry = y + header_h + r * row_h
        if ry < y + h - 2:
            draw.line([x, ry, x + w, ry], fill=(210, 210, 210))


def draw_fields(draw, font, labels, inputs, clip_x=None, y_offset=0):
    """Draw labels and input fields."""
    for item in labels + inputs:
        if clip_x is not None:
            item['x'] = max(item['x'], clip_x + 4)
        item['y'] = item['y'] + y_offset

    for lbl in labels:
        # Labels are left-aligned at their own XML position.
        # Therefore left-justifies label text within the label element, and each
        # label already carries the correct PosX from the form designer.
        lines = wrap_text(lbl['caption'], 28)
        line_height = 12
        for i, line in enumerate(lines):
            draw.text((lbl['x'], lbl['y'] + i * line_height), line, fill=BLACK, font=font)

    for inp in inputs:
        if inp['type'] == 'checkbox':
            draw.rectangle(
                [inp['x'], inp['y'], inp['x'] + 12, inp['y'] + 12],
                fill=WHITE, outline=BORDER_COLOR,
            )
            if inp['caption']:
                draw.text((inp['x'] + 16, inp['y']), inp['caption'], fill=BLACK, font=font)

        elif inp['type'] == 10:
            draw_table_field(draw, font, inp)

        elif inp['w'] > 150 and inp['h'] > 40:
            # Multi-line text area — any wide, tall field (not just 'Notes')
            draw.rectangle(
                [inp['x'], inp['y'], inp['x'] + inp['w'], inp['y'] + inp['h']],
                fill=WHITE, outline=BORDER_COLOR,
            )
            sb_x = inp['x'] + inp['w'] - 12
            draw.rectangle(
                [sb_x, inp['y'], inp['x'] + inp['w'], inp['y'] + inp['h']],
                outline=BORDER_COLOR,
            )
            draw.polygon(
                [(sb_x + 4, inp['y'] + 6), (sb_x + 8, inp['y'] + 6), (sb_x + 6, inp['y'] + 3)],
                fill=(128, 128, 128),
            )
            draw.polygon(
                [(sb_x + 4, inp['y'] + inp['h'] - 6), (sb_x + 8, inp['y'] + inp['h'] - 6),
                 (sb_x + 6, inp['y'] + inp['h'] - 3)],
                fill=(128, 128, 128),
            )
            thumb_y = inp['y'] + 15
            draw.rectangle(
                [sb_x + 1, thumb_y, inp['x'] + inp['w'] - 1, thumb_y + 20],
                fill=(200, 200, 200), outline=BORDER_COLOR,
            )

        else:
            draw.rectangle(
                [inp['x'], inp['y'], inp['x'] + inp['w'], inp['y'] + inp['h']],
                fill=WHITE, outline=BORDER_COLOR,
            )
            if inp['type'] == 3:
                # Date dropdown arrow
                arrow_x = inp['x'] + inp['w'] - 16
                draw.rectangle(
                    [arrow_x, inp['y'] + 1, inp['x'] + inp['w'] - 1, inp['y'] + inp['h'] - 1],
                    fill=(240, 240, 240), outline=BORDER_COLOR,
                )
                ax = arrow_x + 5
                ay = inp['y'] + inp['h'] // 2 - 1
                draw.polygon([(ax, ay), (ax + 6, ay), (ax + 3, ay + 4)], fill=BLACK)
            if inp['caption'] in LOOKUP_FIELDS:
                btn_x = inp['x'] + inp['w'] + 2
                draw.rectangle(
                    [btn_x, inp['y'], btn_x + 18, inp['y'] + inp['h']],
                    fill=(240, 240, 240), outline=BORDER_COLOR,
                )
                draw.text((btn_x + 3, inp['y'] + 3), "...", fill=BLACK, font=font)


# ---------------------------------------------------------------------------
# Category Rendering
# ---------------------------------------------------------------------------
def safe_filename(name):
    return re.sub(r'[^a-zA-Z0-9_-]', '_', name)


def render_simple_category(name, cat, output_dir, font):
    ctgry_no = get_text(cat, 'CtgryNo') or 'x'
    width  = int(get_text(cat, 'Width')  or 300)
    height = int(get_text(cat, 'Height') or 300)
    img_width  = scale_x(width)  + PADDING_LEFT
    img_height = scale_y(height) + PADDING_TOP + 20
    img_width  = max(img_width,  300)
    img_height = max(img_height, 150)

    img  = Image.new('RGB', (img_width, img_height), BG_COLOR)
    draw = ImageDraw.Draw(img)

    draw.rectangle([4, 4, img_width - 4, img_height - 4], outline=(180, 180, 180), width=1)
    draw.rectangle([5, 5, img_width - 5, 27], fill=(240, 240, 240), outline=(180, 180, 180))
    draw.text((10, 9), name, fill=BLACK, font=font)
    draw.rectangle([img_width - 25, 7, img_width - 7, 25], outline=BORDER_COLOR)
    draw.text((img_width - 19, 9), "X", fill=BLACK, font=font)

    fields = cat.find('Fields')
    if fields is not None:
        labels, inputs = get_fields_for_rendering(fields.findall('Field'))
        draw_fields(draw, font, labels, inputs)

    output_path = os.path.join(output_dir, f'{safe_filename(name)}_{ctgry_no}.png')
    img.save(output_path)
    return output_path


def render_tabbed_category(name, cat, output_dir, font):
    ctgry_no = get_text(cat, 'CtgryNo') or 'x'
    fields = cat.find('Fields')
    if fields is None:
        return []

    tab_control    = None
    tabs           = []
    base_fields    = []
    fields_by_tab  = {}

    for field in fields.findall('Field'):
        type_no = int(get_text(field, 'TypeNo') or 1)
        if type_no == 13:
            tab_control = field
            tab_info = field.find('TabInfo')
            if tab_info is not None:
                tabs_elem = tab_info.find('Tabs')
                if tabs_elem is not None:
                    for t in tabs_elem.findall('T'):
                        tab_no  = get_text(t, 'TabNo')
                        tab_pos = get_text(t, 'TabPos')
                        tcapt   = t.find('TabCapt')
                        tc      = tcapt.find('.//S') if tcapt is not None else None
                        tab_cap = tc.text if tc is not None else f'Tab{tab_no}'
                        tabs.append({
                            'no':      tab_no,
                            'pos':     tab_pos,
                            'caption': tab_cap,
                        })
        else:
            show_in_tab = get_text(field, 'ShowInTabNo')
            if show_in_tab is None:
                base_fields.append(field)
            else:
                fields_by_tab.setdefault(show_in_tab, []).append(field)

    # Show ALL tabs — Visible=0 in Therefore often means "system tab" not "hidden"
    all_tabs = sorted(tabs, key=lambda t: int(t['pos'] or 999))
    if not all_tabs:
        return []

    tc_x = scale_x(get_text(tab_control, 'PosX'))
    tc_y = scale_y(get_text(tab_control, 'PosY'))
    tc_w = scale_w(get_text(tab_control, 'Width'))
    tc_h = scale_h(get_text(tab_control, 'Height'))

    img_width  = max(tc_x + tc_w + PADDING_LEFT, 400)
    img_height = max(tc_y + tc_h + 30, 250)

    all_tab_labels, all_tab_inputs = get_fields_for_rendering(
        [f for tab_fields in fields_by_tab.values() for f in tab_fields]
    )
    if all_tab_inputs:
        max_y = max(inp['y'] + inp['h'] for inp in all_tab_inputs)
        img_height = max(img_height, max_y + 20)

    tab_bar_h = 16
    content_y = tc_y + tab_bar_h

    outputs = []
    for active_tab in all_tabs:
        active_tab_no = active_tab['no']
        img  = Image.new('RGB', (img_width, img_height), BG_COLOR)
        draw = ImageDraw.Draw(img)

        # Window chrome
        draw.rectangle([4, 4, img_width - 4, img_height - 4], outline=(180, 180, 180), width=1)
        draw.rectangle([5, 5, img_width - 5, 27], fill=(240, 240, 240), outline=(180, 180, 180))
        draw.text((10, 9), name, fill=BLACK, font=font)
        draw.rectangle([img_width - 25, 7, img_width - 7, 25], outline=BORDER_COLOR)
        draw.text((img_width - 19, 9), "X", fill=BLACK, font=font)

        # Base fields above tab control
        base_labels, base_inputs = get_fields_for_rendering(base_fields)
        draw_fields(draw, font, base_labels, base_inputs)

        # Tab control outer border
        draw.rectangle([tc_x, tc_y, tc_x + tc_w, tc_y + tc_h], outline=BORDER_COLOR)

        # Tab buttons
        tab_x = tc_x + 2
        for tab in all_tabs:
            is_active = tab['no'] == active_tab_no
            bbox  = draw.textbbox((0, 0), tab['caption'], font=font)
            tab_w = (bbox[2] - bbox[0]) + 16

            if is_active:
                draw.rectangle(
                    [tab_x, tc_y, tab_x + tab_w, tc_y + tab_bar_h],
                    fill=TAB_ACTIVE_BG,
                )
                draw.line([tab_x,          tc_y, tab_x,          tc_y + tab_bar_h - 1], fill=BORDER_COLOR)
                draw.line([tab_x + tab_w,  tc_y, tab_x + tab_w,  tc_y + tab_bar_h - 1], fill=BORDER_COLOR)
                draw.line([tab_x,          tc_y, tab_x + tab_w,  tc_y], fill=BORDER_COLOR)
                draw.text((tab_x + 8, tc_y + 3), tab['caption'], fill=BLACK, font=font)
            else:
                draw.rectangle(
                    [tab_x, tc_y + 2, tab_x + tab_w, tc_y + tab_bar_h],
                    fill=TAB_INACTIVE_BG, outline=BORDER_COLOR,
                )
                draw.text((tab_x + 8, tc_y + 4), tab['caption'], fill=BLACK, font=font)
            tab_x += tab_w - 1

        # Content area
        draw.rectangle(
            [tc_x + 1, content_y, tc_x + tc_w - 1, tc_y + tc_h - 1],
            fill=BG_COLOR,
        )

        # Tab fields
        tab_labels, tab_inputs = get_fields_for_rendering(
            fields_by_tab.get(active_tab_no, [])
        )
        if tab_inputs or tab_labels:
            min_field_y = min(item['y'] for item in tab_inputs + tab_labels)
            y_offset = max(0, (content_y + 4) - min_field_y)
        else:
            y_offset = 0

        draw_fields(draw, font, tab_labels, tab_inputs, clip_x=tc_x, y_offset=y_offset)

        safe_tab   = safe_filename(active_tab['caption'])
        output_path = os.path.join(
            output_dir, f'{safe_filename(name)}_{ctgry_no}_tab_{safe_tab}.png'
        )
        img.save(output_path)
        outputs.append(output_path)

    return outputs


# ---------------------------------------------------------------------------
# Main Entry Point
# ---------------------------------------------------------------------------
def render_all_categories(input_xml, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    tree = ET.parse(input_xml)
    root = tree.getroot()

    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 11)
    except Exception:
        try:
            font = ImageFont.truetype("arial.ttf", 11)
        except Exception:
            font = ImageFont.load_default()

    generated = []
    for cat in root.findall('.//Category'):
        ctgry_no = get_text(cat, 'CtgryNo')
        name = get_name(cat) or f'Category_{ctgry_no}'
        fields = cat.find('Fields')
        has_tabs = False
        if fields is not None:
            for f in fields.findall('Field'):
                if int(get_text(f, 'TypeNo') or 1) == 13:
                    has_tabs = True
                    break

        if has_tabs:
            paths = render_tabbed_category(name, cat, output_dir, font)
            generated.extend(paths)
            print(f'{name} (tabs): {len(paths)} image(s)')
        else:
            path = render_simple_category(name, cat, output_dir, font)
            generated.append(path)
            print(f'{name}: {path}')

    print(f'\nDone. Generated {len(generated)} image(s) in {output_dir}')
    return generated


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    render_all_categories(sys.argv[1], sys.argv[2])
