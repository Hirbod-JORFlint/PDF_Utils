#!/usr/bin/env python3
"""
PDF Dark Mode Converter (v6.0 - Font-Preservation Focus)
======================================================

An expert-grade utility for converting PDFs to dark mode while maximizing
the preservation of original styles, specifically bold and italic variants.

Core Optimizations for Font Preservation:
- Binary Flag Extraction: Analyzes the font descriptor flags to detect 
  Bold, Italic, and Monospaced properties even when font names are obfuscated.
- Exact Metadata Passthrough: Uses 'dict' extraction to capture 'flags', 
  'font', 'size', and 'color' for every text span.
- Font Synthesis: If the original font is not available, the script 
  synthesizes the style using the extracted binary attributes to ensure 
  the visual weight and slant are maintained.
- Selective Recoloring: Preserves non-black text (e.g., links, colored headers)
  to keep the original document's semantic styling.

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

# ==============================================================================
# Configuration & Themes
# ==============================================================================

@dataclass
class Theme:
    name: str
    bg_color: Tuple[int, int, int]    # RGB 0-255
    fg_color: Tuple[int, int, int]    # RGB 0-255
    desc: str

THEMES = {
    'amoled': Theme('AMOLED', (0, 0, 0), (235, 235, 235), 'Pure black background'),
    'sepia': Theme('Sepia Dark', (43, 30, 30), (219, 203, 189), 'Warm dark brown reading mode'),
    'navy': Theme('Navy', (10, 25, 47), (100, 255, 218), 'Deep midnight blue and soft cyan')
}

# ==============================================================================
# Style & Color Analytics
# ==============================================================================

def get_rgb_normalized(color_int: int) -> Tuple[float, float, float]:
    """Converts PyMuPDF color integer to normalized (0.0-1.0) RGB float."""
    if color_int is None: return (0.0, 0.0, 0.0)
    return (
        ((color_int >> 16) & 0xFF) / 255.0,
        ((color_int >> 8) & 0xFF) / 255.0,
        (color_int & 0xFF) / 255.0
    )

def is_standard_text_color(rgb: Tuple[float, float, float], threshold: float = 0.3) -> bool:
    """Detects if a color is standard black/dark text meant for darkening."""
    luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
    return luminance < threshold

def get_font_style(flags: int) -> str:
    """
    Decodes PyMuPDF font flags to reconstruct style.
    Bit 0: Superscript, Bit 1: Italic, Bit 2: Serifed, Bit 3: Monospaced, Bit 4: Bold
    """
    style = ""
    # PyMuPDF font names usually follow 'helv', 'tiro', 'cour' bases
    # If the original font name isn't working, we use these constants + style chars
    is_italic = flags & 2  # Bit 1
    is_bold = flags & 16   # Bit 4
    
    if is_bold and is_italic:
        style = "bi"
    elif is_bold:
        style = "b"
    elif is_italic:
        style = "i"
    return style

# ==============================================================================
# Vector Mode (Ultra-Fidelity Font Handling)
# ==============================================================================

def run_vector_mode(doc_in: fitz.Document, doc_out: fitz.Document, theme: Theme):
    """
    Reconstructs the PDF while preserving Font Weight, Slant, and Color accents.
    """
    bg_norm = [c/255.0 for c in theme.bg_color]
    fg_norm = [c/255.0 for c in theme.fg_color]

    for page in tqdm(doc_in, desc="Processing (Vector Mode)", unit="page"):
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Background
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_norm, fill=bg_norm)
        shape.commit()

        # 2. Image Preservation
        for img_info in page.get_image_info(xrefs=True):
            try:
                xref = img_info['xref']
                if xref == 0: continue
                pix = fitz.Pixmap(doc_in, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                new_page.insert_image(img_info['bbox'], stream=pix.tobytes())
            except: continue

        # 3. High-Fidelity Text Extraction
        # dict extraction is essential for capturing font flags (bold/italic)
        page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        
        for block in page_dict.get("blocks", []):
            if block["type"] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        orig_rgb = get_rgb_normalized(span["color"])
                        final_color = fg_norm if is_standard_text_color(orig_rgb) else orig_rgb
                        
                        # Style reconstruction
                        font_name = span["font"]
                        font_flags = span["flags"]
                        style_suffix = get_font_style(font_flags)
                        
                        try:
                            # We attempt to insert text using the original font name.
                            # PyMuPDF is excellent at matching 'Helvetica-Bold' or 'TimesNewRoman,BoldItalic'
                            new_page.insert_text(
                                span["origin"],
                                span["text"],
                                fontsize=span["size"],
                                fontname=font_name,
                                color=final_color,
                                morph=None
                            )
                        except:
                            # If the specific font name fails, use a base font + reconstructed style
                            base_font = "helv" # Default sans
                            if font_flags & 4: base_font = "tiro" # Serif
                            if font_flags & 8: base_font = "cour" # Mono
                            
                            new_page.insert_text(
                                span["origin"], 
                                span["text"], 
                                fontsize=span["size"], 
                                fontname=f"{base_font}{style_suffix}", 
                                color=final_color
                            )

# ==============================================================================
# Image Mode (Robust Rasterization)
# ==============================================================================

def run_image_mode(doc_in: fitz.Document, doc_out: fitz.Document, theme: Theme, args):
    """Raster-based processing with original image stenciling."""
    theme_bg = np.array(theme.bg_color, dtype=np.uint8)
    theme_fg = np.array(theme.fg_color, dtype=np.uint8)

    for page in tqdm(doc_in, desc="Processing (Image Mode)", unit="page"):
        pix = page.get_pixmap(dpi=args.dpi)
        img_orig = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        gray = ImageOps.invert(img_orig.convert("L"))
        mask = gray.point(lambda p: 255 if p > args.threshold else 0)
        if args.blur > 0: mask = mask.filter(ImageFilter.GaussianBlur(args.blur))
        
        m_arr = np.array(mask).astype(float) / 255.0
        m_arr = np.expand_dims(m_arr, axis=2)
        
        canvas_bg = np.full_like(np.array(img_orig), theme_bg)
        canvas_fg = np.full_like(np.array(img_orig), theme_fg)
        
        processed_img = Image.fromarray((canvas_fg * m_arr + canvas_bg * (1.0 - m_arr)).astype(np.uint8))

        # Restore images onto the dark background
        scale = args.dpi / 72.0
        for info in page.get_image_info():
            b = [int(v * scale) for v in info['bbox']]
            if b[2] > b[0] and b[3] > b[1]:
                crop = img_orig.crop((max(0, b[0]), max(0, b[1]), min(img_orig.width, b[2]), min(img_orig.height, b[3])))
                # Dim slightly for aesthetics
                crop = ImageEnhance.Brightness(crop).enhance(0.85)
                processed_img.paste(crop, (b[0], b[1]))

        out_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        buf = io.BytesIO()
        processed_img.save(buf, format="JPEG", quality=85)
        out_page.insert_image(out_page.rect, stream=buf.getvalue())

# ==============================================================================
# CLI Entry Point
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="PDF Dark Mode Converter v6.0")
    parser.add_argument("input", help="Input PDF path")
    parser.add_argument("output", help="Output PDF path")
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled")
    parser.add_argument("--mode", choices=["image", "vector"], default="vector")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--threshold", type=int, default=128)
    parser.add_argument("--blur", type=float, default=0.5)

    args = parser.parse_args()
    
    try:
        doc_in = fitz.open(args.input)
        doc_out = fitz.open()
        theme = THEMES[args.theme]
        
        print(f"[*] Starting Conversion: {args.input}")
        print(f"[*] Mode: {args.mode.upper()} | Theme: {theme.name}")

        if args.mode == "vector":
            run_vector_mode(doc_in, doc_out, theme)
        else:
            run_image_mode(doc_in, doc_out, theme, args)
            
        doc_out.save(args.output, deflate=True, garbage=4)
        print(f"\n[+] Success! File saved to: {args.output}")

    except Exception as e:
        print(f"\n[!] Error: {e}")
        sys.exit(1)
    finally:
        if 'doc_in' in locals(): doc_in.close()
        if 'doc_out' in locals(): doc_out.close()

if __name__ == "__main__":
    main()
