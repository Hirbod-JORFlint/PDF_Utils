#!/usr/bin/env python3
"""
PDF Dark Mode Converter (v8.0 - Ultra Fidelity)
==============================================
Optimized for font style preservation (Bold/Italic) and image fidelity.

Dependencies:
    pip install pymupdf pillow numpy tqdm
"""

import argparse
import sys
import io
import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageFilter, ImageOps, ImageEnhance
from tqdm import tqdm
from dataclasses import dataclass
from typing import Tuple

@dataclass
class Theme:
    name: str
    bg_color: Tuple[int, int, int]
    fg_color: Tuple[int, int, int]
    desc: str

THEMES = {
    'amoled': Theme('AMOLED', (0, 0, 0), (235, 235, 235), 'Pure black'),
    'sepia': Theme('Sepia Dark', (43, 30, 30), (219, 203, 189), 'Warm dark brown'),
    'navy': Theme('Navy', (10, 25, 47), (100, 255, 218), 'Deep blue/cyan')
}

def get_rgb_normalized(color_int: int) -> Tuple[float, float, float]:
    if color_int is None: return (0.0, 0.0, 0.0)
    return (((color_int >> 16) & 0xFF) / 255.0, ((color_int >> 8) & 0xFF) / 255.0, (color_int & 0xFF) / 255.0)

def is_standard_text_color(rgb: Tuple[float, float, float], threshold: float = 0.3) -> bool:
    """Detects if a color is standard black/gray text meant to be inverted."""
    luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
    return luminance < threshold

def get_fallback_font(flags: int) -> str:
    """
    Decodes font flags to pick a Base-14 font that preserves weight/slant.
    Bit 1: Italic | Bit 2: Serif | Bit 3: Mono | Bit 4: Bold
    """
    is_italic = flags & 2
    is_serif = flags & 4
    is_mono = flags & 8
    is_bold = flags & 16

    base = "cour" if is_mono else ("tiro" if is_serif else "helv")
    style = ("bi" if is_bold and is_italic else ("b" if is_bold else ("i" if is_italic else "")))
    return f"{base}{style}"

def run_vector_mode(doc_in, doc_out, theme):
    bg_norm = [c/255.0 for c in theme.bg_color]
    fg_norm = [c/255.0 for c in theme.fg_color]

    for page in tqdm(doc_in, desc="Vector Processing", unit="page"):
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Background Fill
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_norm, fill=bg_norm)
        shape.commit()

        # 2. Image Preservation (Original Bitstream)
        for img in page.get_image_info(xrefs=True):
            try:
                xref = img['xref']
                if xref > 0:
                    pix = fitz.Pixmap(doc_in, xref)
                    if pix.n - pix.alpha > 3: pix = fitz.Pixmap(fitz.csRGB, pix)
                    new_page.insert_image(img['bbox'], stream=pix.tobytes())
            except: continue

        # 3. High-Fidelity Text Reconstruction
        # 
        page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        
        for block in page_dict.get("blocks", []):
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        orig_rgb = get_rgb_normalized(span["color"])
                        final_color = fg_norm if is_standard_text_color(orig_rgb) else orig_rgb
                        
                        try:
                            # Primary: Original Font
                            new_page.insert_text(
                                span["origin"], span["text"],
                                fontsize=span["size"], fontname=span["font"],
                                color=final_color
                            )
                        except Exception:
                            # Fallback: Flag-aware synthesis (Handles Bold/Italic)
                            fallback = get_fallback_font(span["flags"])
                            new_page.insert_text(
                                span["origin"], span["text"],
                                fontsize=span["size"], fontname=fallback,
                                color=final_color
                            )

def run_image_mode(doc_in, doc_out, theme, args):
    """Robust image mode with smart stenciling to keep photos original."""
    theme_bg = np.array(theme.bg_color, dtype=np.uint8)
    theme_fg = np.array(theme.fg_color, dtype=np.uint8)

    for page in tqdm(doc_in, desc="Image Processing", unit="page"):
        pix = page.get_pixmap(dpi=args.dpi)
        img_orig = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Create Luminance Mask
        gray = ImageOps.invert(img_orig.convert("L"))
        mask = gray.point(lambda p: 255 if p > args.threshold else 0)
        if args.blur > 0: mask = mask.filter(ImageFilter.GaussianBlur(args.blur))
        
        m_arr = np.expand_dims(np.array(mask).astype(float) / 255.0, axis=2)
        canvas_bg, canvas_fg = np.full_like(np.array(img_orig), theme_bg), np.full_like(np.array(img_orig), theme_fg)
        processed_img = Image.fromarray((canvas_fg * m_arr + canvas_bg * (1.0 - m_arr)).astype(np.uint8))

        # Stencil original images back in
        scale = args.dpi / 72.0
        for info in page.get_image_info():
            b = [int(v * scale) for v in info['bbox']]
            if b[2] > b[0] and b[3] > b[1]:
                crop = img_orig.crop((max(0, b[0]), max(0, b[1]), min(img_orig.width, b[2]), min(img_orig.height, b[3])))
                processed_img.paste(ImageEnhance.Brightness(crop).enhance(0.85), (b[0], b[1]))

        out_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        buf = io.BytesIO(); processed_img.save(buf, format="JPEG", quality=85)
        out_page.insert_image(out_page.rect, stream=buf.getvalue())

def main():
    parser = argparse.ArgumentParser(description="PDF Dark Mode Converter v8.0")
    parser.add_argument("input", help="Input PDF path"); parser.add_argument("output", help="Output PDF path")
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled")
    parser.add_argument("--mode", choices=["image", "vector"], default="vector")
    parser.add_argument("--dpi", type=int, default=150); parser.add_argument("--threshold", type=int, default=128)
    parser.add_argument("--blur", type=float, default=0.5)

    args = parser.parse_args()
    try:
        doc_in, doc_out = fitz.open(args.input), fitz.open()
        if args.mode == "vector": run_vector_mode(doc_in, doc_out, THEMES[args.theme])
        else: run_image_mode(doc_in, doc_out, THEMES[args.theme], args)
        doc_out.save(args.output, deflate=True, garbage=4)
        print(f"\n[+] Saved: {args.output}")
    except Exception as e:
        print(f"\n[!] Error: {e}"); sys.exit(1)
    finally:
        if 'doc_in' in locals(): doc_in.close()
        if 'doc_out' in locals(): doc_out.close()

if __name__ == "__main__":
    main()
