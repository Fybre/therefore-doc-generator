#!/usr/bin/env python3
"""
themes.py — Template-based styling for Therefore documentation output.

Load a YAML theme file or fall back to the built-in default that replicates
the original hard-coded colours and fonts.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx.shared import RGBColor


@dataclass
class Theme:
    """Colour / font / size theme for both Word documents and image renders."""

    # ── Colours (hex strings without #) ─────────────────────────────────────
    colors: dict[str, str] = field(default_factory=dict)

    # ── Fonts ───────────────────────────────────────────────────────────────
    fonts: dict[str, str] = field(default_factory=dict)

    # ── Heading sizes (pt) ──────────────────────────────────────────────────
    heading_sizes: dict[str, int] = field(default_factory=dict)

    # ── Renderer-specific colours (PIL RGB tuples) ──────────────────────────
    renderer: dict[str, tuple[int, int, int]] = field(default_factory=dict)

    # ── Mermaid diagram colours ─────────────────────────────────────────────
    mermaid: dict[str, dict[str, str]] = field(default_factory=dict)

    # -----------------------------------------------------------------------
    # Helpers
    # -----------------------------------------------------------------------
    def hex(self, name: str) -> str:
        """Return a 6-char hex colour string (no #)."""
        return self.colors.get(name, "000000")

    def rgb(self, name: str) -> tuple[int, int, int]:
        """Return an (r, g, b) tuple for PIL."""
        h = self.hex(name)
        return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))

    def docx_rgb(self, name: str) -> RGBColor:
        """Return an RGBColor for python-docx."""
        return RGBColor(*self.rgb(name))

    def font(self, name: str) -> str:
        return self.fonts.get(name, "Calibri")

    def heading_pt(self, name: str) -> int:
        return self.heading_sizes.get(name, 10)

    def renderer_rgb(self, name: str) -> tuple[int, int, int]:
        return self.renderer.get(name, (0, 0, 0))


# ---------------------------------------------------------------------------
# Default theme — matches the original hard-coded values exactly
# ---------------------------------------------------------------------------
DEFAULT_THEME = Theme(
    colors={
        "primary": "1F4E79",
        "grey": "F2F2F2",
        "white": "FFFFFF",
        "black": "000000",
        "border": "888888",
        "accent": "28a745",
        "danger": "dc3545",
        "warning": "856404",
        "info": "004085",
        "muted": "666666",
        "light_muted": "888888",
        "code_bg": "F4F4F4",
        "code_border": "AAAAAA",
        "footer_border": "CCCCCC",
        "shading": "F4F4F4",
        "table_alt": "F2F2F2",
        "table_border": "888888",
        "h1_border": "1F4E79",
        "para_border": "AAAAAA",
        "render_note": "888888",
        "description": "444444",
        "text_dark": "333333",
        "footer_text": "999999",
    },
    fonts={
        "body": "Calibri",
        "mono": "Consolas",
    },
    heading_sizes={
        "h1": 15,
        "h2": 12,
        "h3": 10,
        "title": 22,
        "subtitle": 14,
    },
    renderer={
        "bg": (240, 240, 240),
        "white": (255, 255, 255),
        "black": (0, 0, 0),
        "border": (172, 172, 172),
        "header": (200, 200, 200),
        "tab_active": (240, 240, 240),
        "tab_inactive": (220, 220, 220),
        "chrome_bg": (240, 240, 240),
        "chrome_border": (180, 180, 180),
        "scroll_thumb": (200, 200, 200),
        "scroll_arrow": (128, 128, 128),
        "date_bg": (240, 240, 240),
        "lookup_btn": (240, 240, 240),
        "panel_header": (31, 78, 121),
        "panel_header_text": (255, 255, 255),
        "panel_bg": (248, 250, 252),
        "panel_border": (180, 200, 220),
        "input_bg": (255, 255, 255),
        "input_border": (180, 180, 180),
        "input_placeholder": (190, 190, 190),
        "label": (50, 50, 50),
        "button_primary": (31, 78, 121),
        "button_secondary": (108, 117, 125),
        "button_text": (255, 255, 255),
        "checkbox_tick": (31, 78, 121),
        "grid_header": (220, 230, 240),
        "grid_border": (180, 180, 180),
        "page_separator": (220, 220, 220),
        "page_title_bg": (240, 244, 248),
        "page_title_fg": (31, 78, 121),
        "html_bg": (248, 248, 200),
        "html_border": (200, 200, 100),
        "html_text": (100, 100, 50),
        "required": (200, 40, 40),
    },
    mermaid={
        "startNode": {"fill": "d4edda", "stroke": "28a745"},
        "endNode": {"fill": "f8d7da", "stroke": "dc3545"},
        "manualNode": {"fill": "cce5ff", "stroke": "004085"},
        "autoNode": {"fill": "fff3cd", "stroke": "856404"},
        "catNode": {"fill": "D6E4F0", "stroke": "1F4E79"},
        "kwNode": {"fill": "E8D6F0", "stroke": "6C3483"},
        "efNode": {"fill": "D6F0E8", "stroke": "1E8449"},
        "dtNode": {"fill": "F0E8D6", "stroke": "B7950B"},
    },
)


# ---------------------------------------------------------------------------
# Load / merge
# ---------------------------------------------------------------------------
def load_theme(path: str | os.PathLike | None = None) -> Theme:
    """
    Load a theme from a YAML file.  Missing keys fall back to DEFAULT_THEME.
    Returns DEFAULT_THEME when *path* is None or the file does not exist.
    """
    if path is None:
        return DEFAULT_THEME

    p = Path(path)
    if not p.exists():
        return DEFAULT_THEME

    try:
        import yaml
    except ImportError as exc:
        raise ImportError(
            "PyYAML is required for theme files. Install: pip install PyYAML"
        ) from exc

    raw = yaml.safe_load(p.read_text(encoding="utf-8")) or {}

    # Deep-merge each top-level section
    def _merge_dict(base: dict, override: dict) -> dict:
        merged = dict(base)
        for k, v in override.items():
            if isinstance(v, dict) and isinstance(merged.get(k), dict):
                merged[k] = _merge_dict(merged[k], v)
            else:
                merged[k] = v
        return merged

    colors = _merge_dict(DEFAULT_THEME.colors, raw.get("colors", {}))
    fonts = _merge_dict(DEFAULT_THEME.fonts, raw.get("fonts", {}))
    heading_sizes = _merge_dict(DEFAULT_THEME.heading_sizes, raw.get("heading_sizes", {}))

    # Renderer colours in YAML can be hex strings; convert to RGB tuples
    renderer = dict(DEFAULT_THEME.renderer)
    for k, v in raw.get("renderer", {}).items():
        if isinstance(v, str):
            renderer[k] = tuple(int(v[i : i + 2], 16) for i in (0, 2, 4))
        elif isinstance(v, (list, tuple)) and len(v) == 3:
            renderer[k] = tuple(int(x) for x in v)

    mermaid = _merge_dict(DEFAULT_THEME.mermaid, raw.get("mermaid", {}))

    return Theme(
        colors=colors,
        fonts=fonts,
        heading_sizes=heading_sizes,
        renderer=renderer,
        mermaid=mermaid,
    )
