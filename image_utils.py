"""
image_utils.py — Shared helpers for rendered images.
"""
from PIL import Image


def save_png(img, path: str, colors: int = 256) -> None:
    """
    Save `img` as a palette PNG with at most `colors` colours.

    Rendered forms and diagrams use a handful of flat colours, so a palette cuts
    the file to roughly a third of a full-colour PNG with no visible change.
    """
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA").quantize(colors, method=Image.Quantize.FASTOCTREE)
    else:
        img = img.convert("RGB").quantize(colors)
    img.save(path, "PNG", optimize=True)
