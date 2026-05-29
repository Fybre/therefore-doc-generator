#!/usr/bin/env python3
"""
Therefore Category Index Form Renderer — SVG output

Reuses the layout logic from render_categories.py but produces scalable
SVG images instead of PNGs.  Text remains selectable and the files scale
cleanly in Word and browsers.

Usage:
    python render_categories_svg.py <input_xml> <output_dir>
"""

import html as _html
import os
import re
import sys
import xml.etree.ElementTree as ET

from themes import DEFAULT_THEME, Theme

# Import pure layout helpers from the PNG renderer
from render_categories import (
    get_caption,
    get_fields_for_rendering,
    get_name,
    get_table_columns,
    get_text,
    is_valid_field,
    safe_filename,
    scale_h,
    scale_w,
    scale_x,
    scale_y,
    wrap_text,
)


# ---------------------------------------------------------------------------
# SVG drawing abstraction (mimics PIL ImageDraw interface)
# ---------------------------------------------------------------------------
class SVGDraw:
    def __init__(self, width: int, height: int, theme: Theme = None):
        self.width = width
        self.height = height
        self._theme = theme or DEFAULT_THEME
        self._buf: list[str] = []
        self._font_size = 22

    # -- helpers --------------------------------------------------------------
    def _colour(self, c) -> str:
        if isinstance(c, tuple) and len(c) == 3:
            return f"rgb({c[0]},{c[1]},{c[2]})"
        if isinstance(c, str) and not c.startswith("#"):
            return f"#{c}"
        return str(c)

    def _txt_y(self, y: int) -> int:
        """PIL places text by top-left; SVG <text> uses baseline."""
        return y + self._font_size - 1

    # -- primitives -----------------------------------------------------------
    def rectangle(self, coords, fill=None, outline=None, width=1):
        x0, y0, x1, y1 = coords
        w, h = x1 - x0, y1 - y0
        attrs = [f'x="{x0}"', f'y="{y0}"', f'width="{w}"', f'height="{h}"']
        if fill is not None:
            attrs.append(f'fill="{self._colour(fill)}"')
        if outline is not None:
            attrs.append(f'stroke="{self._colour(outline)}"')
            attrs.append(f'stroke-width="{width}"')
        self._buf.append(f"  <rect {' '.join(attrs)}/>")

    def rounded_rectangle(self, coords, radius=4, fill=None, outline=None, width=1):
        x0, y0, x1, y1 = coords
        w, h = x1 - x0, y1 - y0
        attrs = [f'x="{x0}"', f'y="{y0}"', f'width="{w}"', f'height="{h}"', f'rx="{radius}"']
        if fill is not None:
            attrs.append(f'fill="{self._colour(fill)}"')
        if outline is not None:
            attrs.append(f'stroke="{self._colour(outline)}"')
            attrs.append(f'stroke-width="{width}"')
        self._buf.append(f"  <rect {' '.join(attrs)}/>")

    def text(self, pos, text, fill=None, font=None):
        x, y = pos
        safe = _html.escape(str(text) if text is not None else "")
        attrs = [f'x="{x}"', f'y="{self._txt_y(y)}"']
        if fill is not None:
            attrs.append(f'fill="{self._colour(fill)}"')
        attrs.append('font-family="Helvetica, Arial, sans-serif"')
        attrs.append(f'font-size="{self._font_size}px"')
        self._buf.append(f"  <text {' '.join(attrs)}>{safe}</text>")

    def line(self, coords, fill=None, width=1):
        x0, y0, x1, y1 = coords
        attrs = [f'x1="{x0}"', f'y1="{y0}"', f'x2="{x1}"', f'y2="{y1}"']
        if fill is not None:
            attrs.append(f'stroke="{self._colour(fill)}"')
        attrs.append(f'stroke-width="{width}"')
        self._buf.append(f"  <line {' '.join(attrs)}/>")

    def polygon(self, points, fill=None, outline=None):
        pts = " ".join(f"{x},{y}" for x, y in points)
        attrs = [f'points="{pts}"']
        if fill is not None:
            attrs.append(f'fill="{self._colour(fill)}"')
        if outline is not None:
            attrs.append(f'stroke="{self._colour(outline)}"')
        self._buf.append(f"  <polygon {' '.join(attrs)}/>")

    def ellipse(self, coords, fill=None, outline=None):
        x0, y0, x1, y1 = coords
        cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
        rx, ry = (x1 - x0) // 2, (y1 - y0) // 2
        attrs = [f'cx="{cx}"', f'cy="{cy}"', f'rx="{rx}"', f'ry="{ry}"']
        if fill is not None:
            attrs.append(f'fill="{self._colour(fill)}"')
        if outline is not None:
            attrs.append(f'stroke="{self._colour(outline)}"')
        self._buf.append(f"  <ellipse {' '.join(attrs)}/>")

    # -- output ---------------------------------------------------------------
    def save(self, path: str):
        svg = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.width}" height="{self.height}">\n'
            f'  <rect x="0" y="0" width="{self.width}" height="{self.height}" '
            f'fill="{self._colour(self._theme.renderer_rgb("bg"))}"/>\n'
            + "\n".join(self._buf)
            + "\n</svg>\n"
        )
        with open(path, "w", encoding="utf-8") as f:
            f.write(svg)


# ---------------------------------------------------------------------------
# Drawing helpers (mirrors render_categories.py)
# ---------------------------------------------------------------------------
def draw_table_field(draw: SVGDraw, inp):
    """Render a TypeNo=10 table field with column header row."""
    x, y, w, h = inp["x"], inp["y"], inp["w"], inp["h"]
    cols = inp.get("columns", [])
    t = draw._theme

    draw.rectangle([x, y, x + w, y + h], fill=t.renderer_rgb("white"), outline=t.renderer_rgb("border"))

    if not cols:
        draw.text((x + 4, y + 4), "(table)", fill=(150, 150, 150))
        return

    header_h = 16
    draw.rectangle([x, y, x + w, y + header_h], fill=t.renderer_rgb("header"), outline=t.renderer_rgb("border"))

    total_w = sum(c["width"] for c in cols)
    col_x = x + 1
    for i, col in enumerate(cols):
        col_w = int((col["width"] / total_w) * (w - 2))
        # Approximate centering — SVG text measurement is tricky without a browser,
        # so we left-align with a small indent.
        draw.text((col_x + 4, y + 3), col["caption"], fill=t.renderer_rgb("black"))
        if i < len(cols) - 1:
            draw.line([col_x + col_w, y, col_x + col_w, y + header_h], fill=t.renderer_rgb("border"))
        col_x += col_w

    row_h = max(16, (h - header_h) // 3)
    for r in range(1, 3):
        ry = y + header_h + r * row_h
        if ry < y + h - 2:
            draw.line([x, ry, x + w, ry], fill=(210, 210, 210))


def _wrap_svg_label(text, max_px, font_size):
    """Word-wrap label text to fit within max_px using font_size as avg char width basis."""
    if not text:
        return []
    avg_char_w = font_size * 0.62
    words = text.split()
    lines, current, current_w = [], "", 0.0
    for word in words:
        word_w = len(word) * avg_char_w
        sep_w = avg_char_w if current else 0.0
        if current and current_w + sep_w + word_w > max_px:
            lines.append(current)
            current, current_w = word, word_w
        else:
            current = (current + " " + word).strip()
            current_w += sep_w + word_w
    if current:
        lines.append(current)
    return lines


def draw_fields(draw: SVGDraw, labels, inputs, clip_x=None, y_offset=0):
    """Draw labels and input fields."""
    t = draw._theme

    for item in labels + inputs:
        if clip_x is not None:
            item["x"] = max(item["x"], clip_x + 4)
        item["y"] = item["y"] + y_offset

    # Draw inputs first so label text renders on top (prevents clipping at field edges)
    for inp in inputs:
        if inp["type"] == "checkbox":
            draw.rectangle(
                [inp["x"], inp["y"], inp["x"] + 12, inp["y"] + 12],
                fill=t.renderer_rgb("white"), outline=t.renderer_rgb("border"),
            )
            if inp["caption"]:
                draw.text((inp["x"] + 16, inp["y"]), inp["caption"], fill=t.renderer_rgb("black"))

        elif inp["type"] == 10:
            draw_table_field(draw, inp)

        elif inp["w"] > 150 and inp["h"] > 40:
            # Multi-line text area
            draw.rectangle(
                [inp["x"], inp["y"], inp["x"] + inp["w"], inp["y"] + inp["h"]],
                fill=t.renderer_rgb("white"), outline=t.renderer_rgb("border"),
            )
            sb_x = inp["x"] + inp["w"] - 12
            draw.rectangle(
                [sb_x, inp["y"], inp["x"] + inp["w"], inp["y"] + inp["h"]],
                outline=t.renderer_rgb("border"),
            )
            draw.polygon(
                [(sb_x + 4, inp["y"] + 6), (sb_x + 8, inp["y"] + 6), (sb_x + 6, inp["y"] + 3)],
                fill=t.renderer_rgb("scroll_arrow"),
            )
            draw.polygon(
                [(sb_x + 4, inp["y"] + inp["h"] - 6), (sb_x + 8, inp["y"] + inp["h"] - 6),
                 (sb_x + 6, inp["y"] + inp["h"] - 3)],
                fill=t.renderer_rgb("scroll_arrow"),
            )
            thumb_y = inp["y"] + 15
            draw.rectangle(
                [sb_x + 1, thumb_y, inp["x"] + inp["w"] - 1, thumb_y + 20],
                fill=t.renderer_rgb("scroll_thumb"), outline=t.renderer_rgb("border"),
            )

        else:
            draw.rectangle(
                [inp["x"], inp["y"], inp["x"] + inp["w"], inp["y"] + inp["h"]],
                fill=t.renderer_rgb("white"), outline=t.renderer_rgb("border"),
            )

            # Dropdown arrow (date fields, and fields explicitly marked as dropdowns)
            if inp["type"] == 3 or inp.get("dropdown"):
                arrow_x = inp["x"] + inp["w"] - 18
                draw.rectangle(
                    [arrow_x, inp["y"] + 1, inp["x"] + inp["w"] - 1, inp["y"] + inp["h"] - 1],
                    fill=t.renderer_rgb("date_bg"), outline=t.renderer_rgb("border"),
                )
                ax = arrow_x + 6
                ay = inp["y"] + inp["h"] // 2 - 1
                draw.polygon([(ax - 3, ay), (ax + 3, ay), (ax, ay + 4)], fill=t.renderer_rgb("black"))

            # Lookup button
            from render_categories import LOOKUP_FIELDS
            if inp["caption"] in LOOKUP_FIELDS:
                btn_x = inp["x"] + inp["w"] + 2
                draw.rectangle(
                    [btn_x, inp["y"], btn_x + 18, inp["y"] + inp["h"]],
                    fill=t.renderer_rgb("lookup_btn"), outline=t.renderer_rgb("border"),
                )
                draw.text((btn_x + 3, inp["y"] + 3), "...", fill=t.renderer_rgb("black"))

    # Draw labels last so text sits on top of input field borders
    for lbl in labels:
        max_px = lbl["w"] * 1.20
        lines = _wrap_svg_label(lbl["caption"], max_px, draw._font_size)
        line_height = 32  # 16 logical px * SS=2
        for i, line in enumerate(lines):
            draw.text((lbl["x"], lbl["y"] + i * line_height), line, fill=t.renderer_rgb("black"))


# ---------------------------------------------------------------------------
# Category rendering
# ---------------------------------------------------------------------------
def render_simple_category(name, cat, output_dir, theme=None):
    ctgry_no = get_text(cat, "CtgryNo") or "x"
    width = int(get_text(cat, "Width") or 300)
    height = int(get_text(cat, "Height") or 300)
    img_width = scale_x(width) + 12
    img_height = scale_y(height) + 35 + 20
    img_width = max(img_width, 300)
    img_height = max(img_height, 150)

    t = theme or DEFAULT_THEME
    draw = SVGDraw(img_width, img_height, theme=t)

    draw.rectangle([4, 4, img_width - 4, img_height - 4], outline=t.renderer_rgb("chrome_border"))
    draw.rectangle([5, 5, img_width - 5, 27], fill=t.renderer_rgb("chrome_bg"), outline=t.renderer_rgb("chrome_border"))
    draw.text((10, 9), name, fill=t.renderer_rgb("black"))
    draw.rectangle([img_width - 25, 7, img_width - 7, 25], outline=t.renderer_rgb("border"))
    draw.text((img_width - 19, 9), "X", fill=t.renderer_rgb("black"))

    fields = cat.find("Fields")
    if fields is not None:
        labels, inputs = get_fields_for_rendering(fields.findall("Field"))
        draw_fields(draw, labels, inputs)

    output_path = os.path.join(output_dir, f"{safe_filename(name)}_{ctgry_no}.svg")
    draw.save(output_path)
    return output_path


def render_tabbed_category(name, cat, output_dir, theme=None):
    ctgry_no = get_text(cat, "CtgryNo") or "x"
    fields = cat.find("Fields")
    if fields is None:
        return []

    tab_control = None
    tabs = []
    base_fields = []
    fields_by_tab = {}

    for field in fields.findall("Field"):
        type_no = int(get_text(field, "TypeNo") or 1)
        if type_no == 13:
            tab_control = field
            tab_info = field.find("TabInfo")
            if tab_info is not None:
                tabs_elem = tab_info.find("Tabs")
                if tabs_elem is not None:
                    for t_el in tabs_elem.findall("T"):
                        tab_no = get_text(t_el, "TabNo")
                        tab_pos = get_text(t_el, "TabPos")
                        tcapt = t_el.find("TabCapt")
                        tc = tcapt.find(".//S") if tcapt is not None else None
                        tab_cap = tc.text if tc is not None else f"Tab{tab_no}"
                        tabs.append({"no": tab_no, "pos": tab_pos, "caption": tab_cap})
        else:
            show_in_tab = get_text(field, "ShowInTabNo")
            if show_in_tab is None:
                base_fields.append(field)
            else:
                fields_by_tab.setdefault(show_in_tab, []).append(field)

    all_tabs = sorted(tabs, key=lambda t: int(t["pos"] or 999))
    if not all_tabs:
        return []

    tc_x = scale_x(get_text(tab_control, "PosX"))
    tc_y = scale_y(get_text(tab_control, "PosY"))
    tc_w = scale_w(get_text(tab_control, "Width"))
    tc_h = scale_h(get_text(tab_control, "Height"))

    img_width = max(tc_x + tc_w + 12, 400)
    img_height = max(tc_y + tc_h + 30, 250)

    all_tab_labels, all_tab_inputs = get_fields_for_rendering(
        [f for tab_fields in fields_by_tab.values() for f in tab_fields]
    )
    if all_tab_inputs:
        max_y = max(inp["y"] + inp["h"] for inp in all_tab_inputs)
        img_height = max(img_height, max_y + 20)

    tab_bar_h = 16
    content_y = tc_y + tab_bar_h

    t = theme or DEFAULT_THEME
    outputs = []
    for active_tab in all_tabs:
        active_tab_no = active_tab["no"]
        draw = SVGDraw(img_width, img_height, theme=t)

        # Window chrome
        draw.rectangle([4, 4, img_width - 4, img_height - 4], outline=t.renderer_rgb("chrome_border"))
        draw.rectangle([5, 5, img_width - 5, 27], fill=t.renderer_rgb("chrome_bg"), outline=t.renderer_rgb("chrome_border"))
        draw.text((10, 9), name, fill=t.renderer_rgb("black"))
        draw.rectangle([img_width - 25, 7, img_width - 7, 25], outline=t.renderer_rgb("border"))
        draw.text((img_width - 19, 9), "X", fill=t.renderer_rgb("black"))

        # Base fields above tab control
        base_labels, base_inputs = get_fields_for_rendering(base_fields)
        draw_fields(draw, base_labels, base_inputs)

        # Tab control outer border
        draw.rectangle([tc_x, tc_y, tc_x + tc_w, tc_y + tc_h], outline=t.renderer_rgb("border"))

        # Tab buttons
        tab_x = tc_x + 2
        for tab in all_tabs:
            is_active = tab["no"] == active_tab_no
            # Approximate text width: 6px per char + 16px padding
            tab_w_approx = len(tab["caption"]) * 6 + 16

            if is_active:
                draw.rectangle(
                    [tab_x, tc_y, tab_x + tab_w_approx, tc_y + tab_bar_h],
                    fill=t.renderer_rgb("tab_active"),
                )
                draw.line([tab_x, tc_y, tab_x, tc_y + tab_bar_h - 1], fill=t.renderer_rgb("border"))
                draw.line([tab_x + tab_w_approx, tc_y, tab_x + tab_w_approx, tc_y + tab_bar_h - 1], fill=t.renderer_rgb("border"))
                draw.line([tab_x, tc_y, tab_x + tab_w_approx, tc_y], fill=t.renderer_rgb("border"))
                draw.text((tab_x + 8, tc_y + 3), tab["caption"], fill=t.renderer_rgb("black"))
            else:
                draw.rectangle(
                    [tab_x, tc_y + 2, tab_x + tab_w_approx, tc_y + tab_bar_h],
                    fill=t.renderer_rgb("tab_inactive"), outline=t.renderer_rgb("border"),
                )
                draw.text((tab_x + 8, tc_y + 4), tab["caption"], fill=t.renderer_rgb("black"))
            tab_x += tab_w_approx - 1

        # Content area
        draw.rectangle(
            [tc_x + 1, content_y, tc_x + tc_w - 1, tc_y + tc_h - 1],
            fill=t.renderer_rgb("bg"),
        )

        # Tab fields
        tab_labels, tab_inputs = get_fields_for_rendering(
            fields_by_tab.get(active_tab_no, [])
        )
        if tab_inputs or tab_labels:
            min_field_y = min(item["y"] for item in tab_inputs + tab_labels)
            y_offset = max(0, (content_y + 4) - min_field_y)
        else:
            y_offset = 0

        draw_fields(draw, tab_labels, tab_inputs, clip_x=tc_x, y_offset=y_offset)

        safe_tab = safe_filename(active_tab["caption"])
        output_path = os.path.join(
            output_dir, f"{safe_filename(name)}_{ctgry_no}_tab_{safe_tab}.svg"
        )
        draw.save(output_path)
        outputs.append(output_path)

    return outputs


# ---------------------------------------------------------------------------
# Batch entry point
# ---------------------------------------------------------------------------
def render_all_categories(input_xml, output_dir, theme=None):
    os.makedirs(output_dir, exist_ok=True)

    tree = ET.parse(input_xml)
    root = tree.getroot()

    generated = []
    for cat in root.findall(".//Category"):
        ctgry_no = get_text(cat, "CtgryNo")
        name = get_name(cat) or f"Category_{ctgry_no}"
        fields = cat.find("Fields")
        has_tabs = False
        if fields is not None:
            for f in fields.findall("Field"):
                if int(get_text(f, "TypeNo") or 1) == 13:
                    has_tabs = True
                    break

        if has_tabs:
            paths = render_tabbed_category(name, cat, output_dir, theme=theme)
            generated.extend(paths)
            print(f"{name} (tabs): {len(paths)} SVG(s)")
        else:
            path = render_simple_category(name, cat, output_dir, theme=theme)
            generated.append(path)
            print(f"{name}: {path}")

    print(f"\nDone. Generated {len(generated)} SVG(s) in {output_dir}")
    return generated


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    render_all_categories(sys.argv[1], sys.argv[2])
