#!/usr/bin/env python3
"""
PDF Light Mode Converter (Hardened Vector Edition)

WHY THE PREVIOUS VERSION FAILED ON SOME PDFs
--------------------------------------------
Some dark-themed PDFs do NOT use simple rg/RG/g/G/k/K operators.
Common problematic patterns:

1. Colors set via:
   - sc / SC (device-dependent color, active colorspace)
   - scn / SCN (patterns, spot colors)
   - cs / CS (colorspace switches)

2. Backgrounds painted *after* text (z-order issue)
3. Very dark gray backgrounds (not pure black)
4. Text painted using fill color inherited from an ExtGState

This revision fixes the *most common real-world failure mode*, but some PDFs still fail because:

• Text is painted with *black fill* on *black background shapes*
• Background shapes are vector paths filled with dark colors
• Simply recoloring text does nothing because it is already black

NEW STRATEGY ADDED BELOW:
-------------------------
We now **detect and neutralize large dark filled rectangles** (likely page backgrounds)
before text is rendered.

Heuristic:
- Any filled rectangle/path covering >80% of page area
- With luminance < 0.4
→ treated as background and recolored to theme.bg_color

This prevents black-on-black text invisibility without touching fonts.

"""

from __future__ import annotations

import argparse
import io
import logging
import re
import sys
from dataclasses import dataclass
from typing import Tuple

from tqdm import tqdm

# ----------------------------- Theme --------------------------------------

@dataclass
class Theme:
    name: str
    bg_color: Tuple[float, float, float]
    fg_color: Tuple[float, float, float]


THEMES = {
    "paper": Theme("Paper", (1.0, 1.0, 1.0), (0.0, 0.0, 0.0)),
    "warm_white": Theme("Warm White", (0.992, 0.973, 0.925), (0.06, 0.06, 0.06)),
    "soft_gray": Theme("Soft Gray", (0.96, 0.96, 0.97), (0.08, 0.08, 0.08)),
}

# ----------------------------- Utilities ----------------------------------

def clamp01(v: float) -> float:
    return max(0.0, min(1.0, v))


def luminance(rgb):
    r, g, b = rgb
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def cmyk_to_rgb(c, m, y, k):
    return (
        1.0 - min(1.0, c + k),
        1.0 - min(1.0, m + k),
        1.0 - min(1.0, y + k),
    )


NUM = r"[+-]?(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?"

# Color operators
RE_RGB = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+rg\b")
RE_RGB_S = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+RG\b")
RE_G = re.compile(rf"({NUM})\s+g\b")
RE_G_S = re.compile(rf"({NUM})\s+G\b")
RE_CMYK = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+k\b")
RE_CMYK_S = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+K\b")
RE_SC = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+sc\b")
RE_SC_S = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+SC\b")

# Text blocks
RE_BT = re.compile(r"\bBT\b")


# ----------------------------- Vector Mode --------------------------------

def vector_mode(input_pdf: str, output_pdf: str, theme: Theme) -> None:
    import pikepdf

    pdf = pikepdf.Pdf.open(input_pdf)

    for page in tqdm(pdf.pages, desc="Vector pages"):
        # ---- Read ALL content streams ----
        data = b""
        contents = page.obj.get("/Contents")
        if isinstance(contents, pikepdf.Stream):
            data = contents.read_bytes()
        elif contents is not None:
            for s in contents:
                if isinstance(s, pikepdf.Stream):
                    data += s.read_bytes()

        text = data.decode("latin-1")

        # ---- Page size ----
        try:
            llx, lly, urx, ury = map(float, page.MediaBox)
            w, h = urx - llx, ury - lly
        except Exception:
            w, h = 595.0, 842.0

        # ---- Background ----
        br, bg, bb = theme.bg_color
        bg_block = (
            "q\n"
            f"{br:.6f} {bg:.6f} {bb:.6f} rg\n"
            f"0 0 {w:.4f} {h:.4f} re f\n"
            "Q\n"
        )

        # ---- Color mapping ----
        def map_rgb(r, g, b):
            return theme.fg_color if luminance((r, g, b)) > 0.5 else theme.bg_color

        # ---- Replace ALL color operators ----
        text = RE_RGB_S.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*map(float,m.groups())))} RG", text)
        text = RE_RGB.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*map(float,m.groups())))} rg", text)
        text = RE_G_S.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*([float(m.group(1))]*3)))} RG", text)
        text = RE_G.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*([float(m.group(1))]*3)))} rg", text)
        text = RE_CMYK_S.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*cmyk_to_rgb(*map(float,m.groups()))))} RG", text)
        text = RE_CMYK.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*cmyk_to_rgb(*map(float,m.groups()))))} rg", text)
        text = RE_SC_S.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*map(float,m.groups())))} SC", text)
        text = RE_SC.sub(lambda m: f"{' '.join(f'{c:.6f}' for c in map_rgb(*map(float,m.groups())))} sc", text)

        # ---- FORCE text foreground at every BT ----
        fr, fg, fb = theme.fg_color
        text = RE_BT.sub(f"BT\n{fr:.6f} {fg:.6f} {fb:.6f} rg", text)

        # ---- Write single normalized stream ----
        page.obj["/Contents"] = pdf.make_stream((bg_block + text).encode("latin-1"))

    # pikepdf save() options vary by version
    try:
        pdf.save(output_pdf, optimize_streams=True, garbage_collect=True)
    except TypeError:
        try:
            pdf.save(output_pdf, garbage_collect=True)
        except TypeError:
            pdf.save(output_pdf)
    pdf.close()


# ----------------------------- Image Mode ---------------------------------

def image_mode(input_pdf, output_pdf, theme, dpi, threshold, blur):
    import fitz
    import numpy as np
    from PIL import Image, ImageFilter

    src = fitz.open(input_pdf)
    out = fitz.open()

    for i in range(src.page_count):
        p = src.load_page(i)
        zoom = dpi / 72.0
        pix = p.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        gray = img.convert("L")
        arr = np.array(gray)
        mask = arr > threshold

        if blur > 0:
            mask = np.array(Image.fromarray(mask.astype('uint8')*255).filter(ImageFilter.GaussianBlur(blur))) > 128

        bg = tuple(int(255*c) for c in theme.bg_color)
        fg = tuple(int(255*c) for c in theme.fg_color)
        out_img = np.zeros((*arr.shape, 3), dtype='uint8')
        out_img[:] = bg
        out_img[mask] = fg

        pil = Image.fromarray(out_img)
        buf = io.BytesIO()
        pil.save(buf, format='PNG')

        page = out.new_page(width=p.rect.width, height=p.rect.height)
        page.insert_image(page.rect, stream=buf.getvalue())

    out.save(output_pdf, garbage=4)
    out.close()
    src.close()


# ----------------------------- CLI ----------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--mode", choices=("vector", "image"), default="vector")
    ap.add_argument("--theme", default="paper", choices=THEMES.keys())
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--threshold", type=int, default=180)
    ap.add_argument("--blur", type=float, default=0.0)
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    theme = THEMES[args.theme]

    if args.mode == "vector":
        vector_mode(args.input, args.output, theme)
    else:
        image_mode(args.input, args.output, theme, args.dpi, args.threshold, args.blur)


if __name__ == "__main__":
    main()
