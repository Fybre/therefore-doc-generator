#!/usr/bin/env python3
"""
Therefore EForm Renderer

Parses the Formio JSON definition stored in an EForm's <FDef> element and
renders a visual mockup as a PNG using PIL.  Produces one image per eform.

Usage:
    python render_eform.py <TheConfiguration.xml> <output_dir>
"""

import html
import json
import os
import re
import sys
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Colours
# ---------------------------------------------------------------------------
BG          = (255, 255, 255)
PANEL_HDR   = (31,  78,  121)   # Therefore blue
PANEL_HDR_T = (255, 255, 255)
PANEL_BG    = (248, 250, 252)
PANEL_BDR   = (180, 200, 220)
INPUT_BG    = (255, 255, 255)
INPUT_BDR   = (180, 180, 180)
INPUT_PH    = (190, 190, 190)
LABEL_FG    = (50,  50,  50)
BTN_BG      = (31,  78,  121)
BTN_TXT     = (255, 255, 255)
BTN_SEC_BG  = (108, 117, 125)
CB_TICK     = (31,  78,  121)
GRID_HDR    = (220, 230, 240)
GRID_BDR    = (180, 180, 180)
PAGE_SEP    = (220, 220, 220)
PAGE_TITLE_BG = (240, 244, 248)
PAGE_TITLE_FG = (31,  78,  121)

# ---------------------------------------------------------------------------
# Layout constants
# ---------------------------------------------------------------------------
CANVAS_W    = 900
PAD         = 24        # outer horizontal padding
COL_W       = CANVAS_W - PAD * 2
LABEL_H     = 18
INPUT_H     = 28
GAP         = 8         # gap below each component
PANEL_PAD   = 12        # inner padding inside panels


def safe_filename(name):
    return re.sub(r"[^a-zA-Z0-9_\-]", "_", name or "unnamed")


# ---------------------------------------------------------------------------
# Font helpers
# ---------------------------------------------------------------------------
def _load_fonts():
    candidates = [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/SFNSDisplay.ttf",
        "/System/Library/Fonts/Arial.ttf",
        "arial.ttf",
    ]
    bold_candidates = [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Arial Bold.ttf",
        "arialbd.ttf",
    ]
    def try_load(paths, size):
        for p in paths:
            try:
                return ImageFont.truetype(p, size)
            except (IOError, OSError):
                pass
        return ImageFont.load_default()

    return {
        "label":   try_load(candidates, 11),
        "input":   try_load(candidates, 11),
        "small":   try_load(candidates, 10),
        "bold":    try_load(bold_candidates, 11),
        "header":  try_load(bold_candidates, 12),
        "page":    try_load(bold_candidates, 13),
        "btn":     try_load(bold_candidates, 11),
    }


FONTS = None


def fonts():
    global FONTS
    if FONTS is None:
        FONTS = _load_fonts()
    return FONTS


# ---------------------------------------------------------------------------
# Draw helpers
# ---------------------------------------------------------------------------
def text_w(draw, text, font):
    try:
        bb = draw.textbbox((0, 0), text, font=font)
        return bb[2] - bb[0]
    except AttributeError:
        return draw.textsize(text, font=font)[0]


def draw_rounded_rect(draw, x0, y0, x1, y1, r, fill, outline=None, width=1):
    draw.rectangle([x0 + r, y0, x1 - r, y1], fill=fill)
    draw.rectangle([x0, y0 + r, x1, y1 - r], fill=fill)
    draw.ellipse([x0, y0, x0 + 2*r, y0 + 2*r], fill=fill)
    draw.ellipse([x1 - 2*r, y0, x1, y0 + 2*r], fill=fill)
    draw.ellipse([x0, y1 - 2*r, x0 + 2*r, y1], fill=fill)
    draw.ellipse([x1 - 2*r, y1 - 2*r, x1, y1], fill=fill)
    if outline:
        draw.rounded_rectangle([x0, y0, x1, y1], radius=r,
                                outline=outline, width=width)


# ---------------------------------------------------------------------------
# Component height estimator (for pre-layout)
# ---------------------------------------------------------------------------
def component_height(comp, avail_w):
    t = comp.get("type", "")
    if comp.get("hidden"):
        return 0
    if t in ("textfield", "number", "datetime", "lookup"):
        return LABEL_H + INPUT_H + GAP
    if t == "textarea":
        return LABEL_H + 72 + GAP
    if t == "select":
        return LABEL_H + INPUT_H + GAP
    if t == "checkbox":
        return 24 + GAP
    if t == "button":
        return 34 + GAP
    if t == "htmlelement":
        content = comp.get("content", "")
        # Skip image-only elements
        if "<img" in content and len(re.sub(r"<[^>]+>", "", content).strip()) < 5:
            return 0
        return 32 + GAP
    if t == "columns":
        cols = comp.get("columns") or []
        if not cols:
            return 0
        col_heights = []
        for col in cols:
            w_frac = (col.get("width") or 6) / 12
            cw = int(avail_w * w_frac) - 8
            col_heights.append(sum(component_height(c, cw)
                                   for c in col.get("components", [])
                                   if not c.get("hidden")))
        return max(col_heights) + GAP if col_heights else GAP
    if t == "panel":
        inner = PAD + sum(component_height(c, avail_w - PANEL_PAD * 2)
                          for c in comp.get("components", [])
                          if not c.get("hidden"))
        return LABEL_H + 6 + inner + PANEL_PAD + GAP
    if t == "datagrid":
        rows = comp.get("defaultValue") or []
        data_rows = max(len(rows), 1)
        return LABEL_H + 28 + data_rows * 24 + GAP
    return 0


def form_height(components, avail_w):
    return sum(component_height(c, avail_w) for c in components
               if not c.get("hidden"))


# ---------------------------------------------------------------------------
# Renderer
# ---------------------------------------------------------------------------
class FormRenderer:
    def __init__(self, draw, x0, avail_w):
        self.draw     = draw
        self.x0       = x0
        self.avail_w  = avail_w
        self.y        = 0
        self.f        = fonts()

    def render_all(self, components):
        for comp in components:
            if comp.get("hidden"):
                continue
            self._render(comp, self.x0, self.avail_w)

    def _render(self, comp, x, w):
        t = comp.get("type", "")
        if t in ("textfield", "number", "lookup"):
            self._input_field(comp, x, w)
        elif t == "textarea":
            self._textarea(comp, x, w)
        elif t == "datetime":
            self._datetime_field(comp, x, w)
        elif t == "select":
            self._select_field(comp, x, w)
        elif t == "checkbox":
            self._checkbox(comp, x, w)
        elif t == "button":
            self._button(comp, x, w)
        elif t == "htmlelement":
            self._html_element(comp, x, w)
        elif t == "columns":
            self._columns(comp, x, w)
        elif t == "panel":
            self._panel(comp, x, w)
        elif t == "datagrid":
            self._datagrid(comp, x, w)

    # -- label helper --------------------------------------------------------
    def _label(self, text, x, required=False):
        if not text:
            return
        lbl = str(text)
        self.draw.text((x, self.y), lbl, font=self.f["label"], fill=LABEL_FG)
        if required:
            tw = text_w(self.draw, lbl, self.f["label"])
            self.draw.text((x + tw + 2, self.y), "*",
                           font=self.f["label"], fill=(200, 40, 40))
        self.y += LABEL_H

    def _input_box(self, x, w, placeholder="", icon=None):
        x1 = x + w
        y1 = self.y + INPUT_H
        self.draw.rectangle([x, self.y, x1, y1],
                            fill=INPUT_BG, outline=INPUT_BDR)
        if placeholder:
            self.draw.text((x + 6, self.y + 7), placeholder,
                           font=self.f["small"], fill=INPUT_PH)
        if icon == "calendar":
            self.draw.text((x1 - 20, self.y + 6), "📅",
                           font=self.f["small"], fill=INPUT_PH)
        elif icon == "search":
            self.draw.text((x1 - 20, self.y + 6), "🔍",
                           font=self.f["small"], fill=INPUT_PH)
        elif icon == "dropdown":
            # Draw a simple triangle
            cx = x1 - 14
            cy = self.y + INPUT_H // 2
            self.draw.polygon([(cx, cy - 3), (cx + 8, cy - 3), (cx + 4, cy + 3)],
                              fill=INPUT_PH)
        self.y += INPUT_H

    # -- component renderers -------------------------------------------------
    def _input_field(self, comp, x, w):
        lbl  = comp.get("label", "")
        req  = comp.get("validate", {}).get("required", False)
        icon = "search" if comp.get("type") == "lookup" else None
        self._label(lbl, x, required=req)
        self._input_box(x, w, icon=icon)
        self.y += GAP

    def _textarea(self, comp, x, w):
        lbl = comp.get("label", "")
        req = comp.get("validate", {}).get("required", False)
        self._label(lbl, x, required=req)
        y1 = self.y + 72
        self.draw.rectangle([x, self.y, x + w, y1],
                            fill=INPUT_BG, outline=INPUT_BDR)
        self.y = y1 + GAP

    def _datetime_field(self, comp, x, w):
        lbl = comp.get("label", "")
        req = comp.get("validate", {}).get("required", False)
        self._label(lbl, x, required=req)
        self._input_box(x, w, icon="calendar")
        self.y += GAP

    def _select_field(self, comp, x, w):
        lbl = comp.get("label", "")
        req = comp.get("validate", {}).get("required", False)
        self._label(lbl, x, required=req)
        # Show first option as placeholder if available
        values = (comp.get("data") or {}).get("values") or []
        ph = values[0].get("label", "") if values else ""
        self._input_box(x, w, placeholder=ph or "Select…", icon="dropdown")
        self.y += GAP

    def _checkbox(self, comp, x, w):
        lbl = comp.get("label", "")
        cx  = x
        cy  = self.y + 4
        self.draw.rectangle([cx, cy, cx + 14, cy + 14],
                            fill=INPUT_BG, outline=INPUT_BDR)
        # draw tick mark
        self.draw.line([(cx + 2, cy + 7), (cx + 5, cy + 11),
                        (cx + 12, cy + 3)], fill=CB_TICK, width=2)
        self.draw.text((cx + 20, self.y + 2), lbl,
                       font=self.f["label"], fill=LABEL_FG)
        self.y += 24 + GAP

    def _button(self, comp, x, w):
        lbl  = comp.get("label", "")
        theme = comp.get("theme", "primary")
        bg = BTN_BG if theme in ("primary", "") else BTN_SEC_BG
        bw   = min(text_w(self.draw, lbl, self.f["btn"]) + 28, w)
        y1   = self.y + 28
        draw_rounded_rect(self.draw, x, self.y, x + bw, y1, 4, bg)
        self.draw.text((x + 10, self.y + 8), lbl,
                       font=self.f["btn"], fill=BTN_TXT)
        self.y = y1 + GAP

    def _html_element(self, comp, x, w):
        content = comp.get("content", "")
        # Skip pure image elements
        if "<img" in content and len(re.sub(r"<[^>]+>", "", content).strip()) < 5:
            return
        # Strip tags, strip style/script blocks, get readable text
        text = re.sub(r"<style[^>]*>.*?</style>", " ", content, flags=re.S)
        text = re.sub(r"<script[^>]*>.*?</script>", " ", text, flags=re.S)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html.unescape(text)
        text = re.sub(r"\s+", " ", text).strip()
        # Skip if it's just CSS/JS noise with no real prose
        if re.match(r"^[\{\}\(\);:#\.\s]+$", text):
            return
        text = text[:90]
        if not text:
            return
        y1 = self.y + 26
        self.draw.rectangle([x, self.y, x + w, y1],
                            fill=(248, 248, 200), outline=(200, 200, 100))
        self.draw.text((x + 6, self.y + 6), text,
                       font=self.f["small"], fill=(100, 100, 50))
        self.y = y1 + GAP

    def _columns(self, comp, x, w):
        cols = comp.get("columns") or []
        if not cols:
            return
        total_units = sum((c.get("width") or 6) for c in cols)
        y_start = self.y
        y_max   = y_start
        cur_x   = x
        for col in cols:
            units  = col.get("width") or 6
            cw     = int(w * units / total_units) - 4
            sub    = col.get("components", [])
            sub    = [c for c in sub if not c.get("hidden")]
            if not sub:
                cur_x += cw + 4
                continue
            self.y = y_start
            sub_r  = FormRenderer(self.draw, cur_x, cw)
            sub_r.y = y_start
            sub_r.render_all(sub)
            y_max  = max(y_max, sub_r.y)
            cur_x += cw + 4
        self.y = y_max + GAP

    def _panel(self, comp, x, w):
        label  = comp.get("label") or comp.get("title") or ""
        children = [c for c in comp.get("components", [])
                    if not c.get("hidden")]
        inner_h = form_height(children, w - PANEL_PAD * 2) + PANEL_PAD

        # Header bar
        hdr_y1 = self.y + LABEL_H + 6
        self.draw.rectangle([x, self.y, x + w, hdr_y1],
                            fill=PANEL_HDR)
        if label:
            self.draw.text((x + 8, self.y + 4), label,
                           font=self.f["header"], fill=PANEL_HDR_T)
        # Body
        body_y0 = hdr_y1
        body_y1 = body_y0 + inner_h
        self.draw.rectangle([x, body_y0, x + w, body_y1],
                            fill=PANEL_BG, outline=PANEL_BDR)
        # Render children
        sub_r = FormRenderer(self.draw, x + PANEL_PAD, w - PANEL_PAD * 2)
        sub_r.y = body_y0 + PANEL_PAD
        sub_r.render_all(children)
        self.y = body_y1 + GAP

    def _datagrid(self, comp, x, w):
        lbl      = comp.get("label", "")
        children = comp.get("components") or []
        self._label(lbl, x)

        if not children:
            self.y += GAP
            return

        col_w = w // len(children) if children else w

        # Header row
        hx = x
        hy = self.y
        for c in children:
            self.draw.rectangle([hx, hy, hx + col_w, hy + 24],
                                fill=GRID_HDR, outline=GRID_BDR)
            self.draw.text((hx + 4, hy + 5), c.get("label", ""),
                           font=self.f["small"], fill=LABEL_FG)
            hx += col_w
        self.y += 24

        # One empty data row
        hx = x
        for _ in children:
            self.draw.rectangle([hx, self.y, hx + col_w, self.y + 24],
                                fill=INPUT_BG, outline=GRID_BDR)
            hx += col_w
        self.y += 24 + GAP


# ---------------------------------------------------------------------------
# Per-page rendering (wizard panels = pages)
# ---------------------------------------------------------------------------
def _is_page(comp):
    return comp.get("type") == "panel"


def render_eform(ef_elem, output_path: str, font=None) -> str | None:
    """
    Render an EForm element to a PNG.
    `font` accepted for API compatibility but unused (we load our own).
    Returns output_path on success, None if no FDef.
    """
    fdef_text = ef_elem.findtext("FDef")
    if not fdef_text:
        return None
    try:
        fdef = json.loads(fdef_text)
    except json.JSONDecodeError:
        return None

    top_components = fdef.get("components") or []
    if not top_components:
        return None

    f = fonts()

    # Separate wizard pages from standalone components
    pages  = [c for c in top_components if _is_page(c) and not c.get("hidden")]
    others = [c for c in top_components if not _is_page(c) and not c.get("hidden")]

    # Estimate total height
    sections = []
    if pages:
        for p in pages:
            sections.append(("page", p))
    # Standalone components go in an implicit section
    if others:
        sections.append(("standalone", others))

    total_h = PAD
    for kind, data in sections:
        if kind == "page":
            children = [c for c in data.get("components", []) if not c.get("hidden")]
            total_h += 30 + PAD  # page title
            total_h += form_height(children, COL_W) + PAD
        else:
            total_h += form_height(data, COL_W) + PAD
    total_h += PAD

    total_h = max(total_h, 200)

    img  = Image.new("RGB", (CANVAS_W, total_h), BG)
    draw = ImageDraw.Draw(img)

    y = PAD

    for kind, data in sections:
        if kind == "page":
            label    = data.get("label") or data.get("title") or ""
            children = [c for c in data.get("components", []) if not c.get("hidden")]
            if not children:
                continue
            # Page title banner
            draw.rectangle([0, y, CANVAS_W, y + 28], fill=PAGE_TITLE_BG)
            draw.text((PAD, y + 6), label, font=f["page"], fill=PAGE_TITLE_FG)
            y += 28 + 8
            r = FormRenderer(draw, PAD, COL_W)
            r.y = y
            r.render_all(children)
            y = r.y + PAD
        else:
            r = FormRenderer(draw, PAD, COL_W)
            r.y = y
            r.render_all(data)
            y = r.y + PAD

    # Crop to actual content
    img = img.crop((0, 0, CANVAS_W, y + PAD))
    img.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Batch entry point
# ---------------------------------------------------------------------------
def render_all_eforms(input_xml: str, output_dir: str) -> list:
    os.makedirs(output_dir, exist_ok=True)
    tree = ET.parse(input_xml)
    root = tree.getroot()

    ef_parent = root.find("EForms")
    if ef_parent is None:
        print("No EForms found.")
        return []

    generated = []
    seen = {}
    for ef in ef_parent:
        name   = ef.findtext("FName") or "EForm"
        fno    = ef.findtext("FNo") or "x"
        safe   = safe_filename(name)
        count  = seen.get(safe, 0)
        seen[safe] = count + 1
        fname  = f"{safe}_{fno}.png"
        path   = os.path.join(output_dir, fname)
        result = render_eform(ef, path)
        if result:
            generated.append(result)
            print(f"  {name}: {path}")
        else:
            print(f"  {name}: skipped (no FDef)")

    print(f"\nDone. {len(generated)} eform image(s) in {output_dir}")
    return generated


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    render_all_eforms(sys.argv[1], sys.argv[2])
