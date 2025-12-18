#!/usr/bin/env python3
"""
PDF Dark Mode Converter (v3.0 - High Fidelity)
==============================================

A professional utility for converting PDFs to dark mode while maximizing
fidelity of fonts, styles, and embedded imagery.

Key Improvements:
- Image Mode: Smart bounding-box detection to preserve original image regions.
- Vector Mode: Intelligent color-swapping that preserves intentional 
  highlight colors (links, accents) while darkening standard text.
- Font Preservation: Uses PyMuPDF's low-level font name mapping.

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
    bg_color: Tuple[int, int, int]  # RGB 0-255
    fg_color: Tuple[int, int, int]  # RGB 0-255

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
    """
    Detects if a color is 'dark' (usually black/gray text on white).
    Luminance-based detection.
    """
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
    """
    Raster-based processing with intelligent image patching to prevent 
    darkening/inverting photos and diagrams.
    """
    bg_col = theme.bg_color
    fg_col = theme.fg_color

    for page in tqdm(doc_in, desc="Rasterizing Pages", unit="page"):
        pix = page.get_pixmap(dpi=args.dpi)
        img_orig = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # 1. Create dark mode base via luminance mask for text/background
        # Invert grayscale so dark text becomes bright (mask)
        gray = ImageOps.invert(img_orig.convert("L"))
        mask = gray.point(lambda p: 255 if p > args.threshold else 0)
        if args.blur > 0:
            mask = mask.filter(ImageFilter.GaussianBlur(args.blur))
        
        mask_arr = np.array(mask).astype(float) / 255.0
        mask_arr = np.expand_dims(mask_arr, axis=2)
        
        base_bg = np.full_like(np.array(img_orig), bg_col)
        base_fg = np.full_like(np.array(img_orig), fg_col)
        
        # Composite: (Foreground * Mask) + (Background * inverse Mask)
        composite = (base_fg * mask_arr + base_bg * (1.0 - mask_arr)).astype(np.uint8)
        processed_img = Image.fromarray(composite)
        
        # 2. Re-paste original images from the PDF to keep them from being inverted
        scale = args.dpi / 72.0
        for img_info in page.get_image_info():
            b = [int(v * scale) for v in img_info['bbox']]
            # Ensure crop is within bounds
            b[0], b[1] = max(0, b[0]), max(0, b[1])
            b[2], b[3] = min(img_orig.width, b[2]), min(img_orig.height, b[3])
            
            if b[2] > b[0] and b[3] > b[1]:
                crop = img_orig.crop((b[0], b[1], b[2], b[3]))
                # Dim them slightly so they aren't blindingly bright in dark mode
                processed_img.paste(dim_image(crop), (b[0], b[1]))

        # 3. Save to output doc
        out_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        img_buffer = io.BytesIO()
        processed_img.save(img_buffer, format="JPEG", quality=90)
        out_page.insert_image(out_page.rect, stream=img_buffer.getvalue())

def run_vector_mode(doc_in, doc_out, theme, args):
    """
    High-fidelity vector reconstruction. Preserves selectable text, 
    font properties, and original images.
    """
    bg_norm = [c/255.0 for c in theme.bg_color]
    fg_norm = [c/255.0 for c in theme.fg_color]

    for page in tqdm(doc_in, desc="Reconstructing Vectors", unit="page"):
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Background Fill
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_norm, fill=bg_norm)
        shape.commit()

        # 2. Extract and Insert Original Images
        for img in page.get_image_info(xrefs=True):
            try:
                xref = img['xref']
                if xref == 0: continue
                pix = fitz.Pixmap(doc_in, xref)
                # Convert to RGB if necessary (e.g., CMYK images)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                
                new_page.insert_image(img['bbox'], stream=pix.tobytes())
            except:
                continue

        # 3. High-Fidelity Text Reconstruction
        # Using specific flags to preserve spacing and ligatures
        dict_data = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        
        for block in dict_data.get("blocks", []):
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        # Style Preservation Logic
                        orig_rgb = get_rgb_from_int(span["color"])
                        
                        # Only darken/swap text if it was originally dark (standard text).
                        # If the text has a distinct color (red, blue, green), keep it for style.
                        final_color = fg_norm if is_dark_color(orig_rgb) else orig_rgb
                        
                        try:
                            # Attempt to use original font names
                            new_page.insert_text(
                                span["origin"],
                                span["text"],
                                fontsize=span["size"],
                                fontname=span["font"],
                                color=final_color
                            )
                        except:
                            # Fallback if specific font cannot be mapped
                            new_page.insert_text(
                                span["origin"], 
                                span["text"], 
                                fontsize=span["size"], 
                                color=final_color
                            )

# ==============================================================================
# CLI Implementation
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(description="Professional PDF Dark Mode Converter")
    parser.add_argument("input", help="Path to input PDF")
    parser.add_argument("output", help="Path to output PDF")
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled", help="Color theme")
    parser.add_argument("--mode", choices=["image", "vector"], default="image", 
                        help="Processing mode: image (robust) or vector (selectable text)")
    parser.add_argument("--dpi", type=int, default=150, help="DPI for image mode")
    parser.add_argument("--threshold", type=int, default=120, help="Text detection threshold (0-255)")
    parser.add_argument("--blur", type=float, default=0.5, help="Mask blur for smoother text edges")

    args = parser.parse_args()
    
    try:
        doc_in = fitz.open(args.input)
        doc_out = fitz.open()
        
        selected_theme = THEMES[args.theme]
        
        if args.mode == "image":
            run_image_mode(doc_in, doc_out, selected_theme, args)
        else:
            run_vector_mode(doc_in, doc_out, selected_theme, args)
            
        # Use deflate and garbage collection to keep file size small
        doc_out.save(args.output, deflate=True, garbage=3)
        print(f"\n[Success] Processed PDF saved to: {args.output}")
        
    except Exception as e:
        print(f"\n[Error] {str(e)}")
        sys.exit(1)
    finally:
        if 'doc_in' in locals(): doc_in.close()
        if 'doc_out' in locals(): doc_out.close()

if __name__ == "__main__":
    main()
