#!/usr/bin/env python3
"""
PDF Dark Mode Converter (v2.0)
==============================

A professional utility for converting PDFs to dark mode while maximizing
fidelity of fonts, styles, and embedded imagery.

Dependencies:
    pip install pymupdf pillow numpy tqdm

Usage:
    python pdf_darkmode.py input.pdf output.pdf --theme navy --mode vector
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
    description: str
    bg_color: Tuple[int, int, int]  # RGB
    fg_color: Tuple[int, int, int]  # RGB

THEMES = {
    'amoled': Theme('AMOLED', 'Pure black background', (0, 0, 0), (235, 235, 235)),
    'sepia': Theme('Sepia Dark', 'Warm dark brown', (43, 30, 30), (219, 203, 189)),
    'navy': Theme('Navy', 'Deep blue and cyan', (10, 25, 47), (100, 255, 218))
}

# ==============================================================================
# Style & Color Helpers
# ==============================================================================

def get_rgb_from_int(color_int: int) -> Tuple[float, float, float]:
    """Converts PyMuPDF color integer to normalized (0-1) RGB."""
    if color_int is None:
        return (0.0, 0.0, 0.0)
    r = ((color_int >> 16) & 0xFF) / 255.0
    g = ((color_int >> 8) & 0xFF) / 255.0
    b = (color_int & 0xFF) / 255.0
    return (r, g, b)

def is_dark_color(rgb: Tuple[float, float, float], threshold: float = 0.4) -> bool:
    """Detects if a color is 'dark' (usually black text on white)."""
    luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
    return luminance < threshold

def dim_image(pil_img: Image.Image, factor: float = 0.8) -> Image.Image:
    """Slightly dims images to reduce eye strain against dark backgrounds."""
    enhancer = ImageEnhance.Brightness(pil_img)
    return enhancer.enhance(factor)

# ==============================================================================
# Processing Modes
# ==============================================================================

def run_image_mode(doc_in, doc_out, theme, args):
    """Raster-based processing with intelligent image patching."""
    bg_col = theme.bg_color
    fg_col = theme.fg_color

    for page in tqdm(doc_in, desc="Rasterizing Pages", unit="page"):
        pix = page.get_pixmap(dpi=args.dpi)
        img_orig = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Create dark mode base via luminance mask
        gray = ImageOps.invert(img_orig.convert("L"))
        mask = gray.point(lambda p: 255 if p > args.threshold else 0)
        if args.blur > 0:
            mask = mask.filter(ImageFilter.GaussianBlur(args.blur))
        
        mask_arr = np.array(mask).astype(float) / 255.0
        mask_arr = np.expand_dims(mask_arr, axis=2)
        
        base_bg = np.full_like(np.array(img_orig), bg_col)
        base_fg = np.full_like(np.array(img_orig), fg_col)
        
        # Composite text over background
        composite = (base_fg * mask_arr + base_bg * (1.0 - mask_arr)).astype(np.uint8)
        processed_img = Image.fromarray(composite)
        
        # Re-paste original images to keep them from being 'inverted/darkened'
        scale = args.dpi / 72.0
        for img_info in page.get_image_info():
            b = [int(v * scale) for v in img_info['bbox']]
            # Crop, dim, and paste
            crop = img_orig.crop((b[0], b[1], b[2], b[3]))
            processed_img.paste(dim_image(crop), (b[0], b[1]))

        # Save to new PDF
        out_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        img_buffer = io.BytesIO()
        processed_img.save(img_buffer, format="JPEG", quality=90)
        out_page.insert_image(out_page.rect, stream=img_buffer.getvalue())

def run_vector_mode(doc_in, doc_out, theme, args):
    """High-fidelity vector reconstruction with image & font preservation."""
    bg_norm = [c/255.0 for c in theme.bg_color]
    fg_norm = [c/255.0 for c in theme.fg_color]

    for page in tqdm(doc_in, desc="Reconstructing Vectors", unit="page"):
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Background
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_norm, fill=bg_norm)
        shape.commit()

        # 2. Preserve Images (Original Fidelity)
        for img in page.get_image_info(xrefs=True):
            try:
                xref = img['xref']
                if xref == 0: continue
                pix = fitz.Pixmap(doc_in, xref)
                # If pixmap is CMYK or similar, convert to RGB
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                
                img_data = pix.tobytes()
                new_page.insert_image(img['bbox'], stream=img_data)
            except Exception:
                continue

        # 3. High-Fidelity Text Reconstruction
        dict_data = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        
        for block in dict_data.get("blocks", []):
            if block["type"] == 0:  # Text
                for line in block["lines"]:
                    for span in line["spans"]:
                        # Extract exact original style
                        orig_rgb = get_rgb_from_int(span["color"])
                        
                        # Decide color: if it was originally dark/black, use theme FG.
                        # If it was colored (blue links, red labels), keep original color.
                        final_color = fg_norm if is_dark_color(orig_rgb) else orig_rgb
                        
                        try:
                            # Try to use original font name to keep style
                            new_page.insert_text(
                                span["origin"],
                                span["text"],
                                fontsize=span["size"],
                                fontname=span["font"],
                                color=final_color,
                                fontfile=None # PyMuPDF tries to match system/standard fonts
                            )
                        except:
                            # Generic fallback for rare font names
                            new_page.insert_text(span["origin"], span["text"], 
                                                 fontsize=span["size"], color=final_color)

# ==============================================================================
# CLI Entry Point
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="Professional PDF Dark Mode Converter")
    parser.add_argument("input", help="Path to input PDF")
    parser.add_argument("output", help="Path to output PDF")
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled")
    parser.add_argument("--mode", choices=["image", "vector"], default="image")
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--threshold", type=int, default=120, help="Text detection sensitivity")
    parser.add_argument("--blur", type=float, default=0.5)

    args = parser.parse_args()
    
    try:
        doc_in = fitz.open(args.input)
        doc_out = fitz.open()
        
        selected_theme = THEMES[args.theme]
        
        if args.mode == "image":
            run_image_mode(doc_in, doc_out, selected_theme, args)
        else:
            run_vector_mode(doc_in, doc_out, selected_theme, args)
            
        doc_out.save(args.output, deflate=True, garbage=3)
        print(f"\n[Done] Dark-mode PDF saved to: {args.output}")
        
    except Exception as e:
        print(f"\n[Error] {str(e)}")
        sys.exit(1)
    finally:
        if 'doc_in' in locals(): doc_in.close()
        if 'doc_out' in locals(): doc_out.close()

if __name__ == "__main__":
    main()
