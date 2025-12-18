#!/usr/bin/env python3
"""
PDF Dark Mode Converter
=======================

A robust tool to convert PDF documents into dark mode variants using 
advanced image processing or vector reconstruction.

Dependencies:
    pip install pymupdf pillow numpy tqdm

Usage:
    python pdf_darkmode.py input.pdf output.pdf --theme navy --mode image
"""

import argparse
import sys
import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageFilter, ImageOps, ImageEnhance
from tqdm import tqdm
from dataclasses import dataclass
from typing import Tuple, List, Dict
import io

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
    'amoled': Theme(
        name='AMOLED',
        description='Pure black background, near-white text (High Contrast)',
        bg_color=(0, 0, 0),
        fg_color=(235, 235, 235)
    ),
    'sepia': Theme(
        name='Sepia Dark',
        description='Dark brown background, warm cream text (Reading Mode)',
        bg_color=(43, 30, 30),
        fg_color=(219, 203, 189)
    ),
    'navy': Theme(
        name='Navy',
        description='Deep blue background, soft cyan text (Night Mode)',
        bg_color=(10, 25, 47),
        fg_color=(100, 255, 218)
    )
}

# ==============================================================================
# Helper Functions
# ==============================================================================

def int_to_rgb(color_int: int) -> Tuple[int, int, int]:
    """Converts PyMuPDF integer color to RGB tuple."""
    r = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    b = color_int & 0xFF
    return (r, g, b)

def is_dark_color(rgb: Tuple[int, int, int], threshold: int = 100) -> bool:
    """Determines if a color is 'dark' (likely standard text)."""
    return (sum(rgb) / 3) < threshold

def enhance_image_for_dark_mode(pil_image: Image.Image) -> Image.Image:
    """
    Slightly dims an image so it doesn't blind the user on a dark background,
    without making it muddy or inverted.
    """
    # 1. Reduce brightness slightly (85%)
    enhancer = ImageEnhance.Brightness(pil_image)
    img = enhancer.enhance(0.85)
    # 2. Slight contrast boost to pop against black
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(1.1)
    return img

# ==============================================================================
# Image Mode Logic (Raster + Smart Image Patching)
# ==============================================================================

def process_page_image(
    page: fitz.Page, 
    theme: Theme, 
    dpi: int, 
    threshold_val: int, 
    blur_radius: float
) -> Image.Image:
    """
    Renders page to image, recolors text, but PRESERVES original images/graphs.
    """
    # 1. Render complete page to high-res image
    pix = page.get_pixmap(dpi=dpi)
    img_original = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # 2. Generate Dark Mode Base (Text Recoloring)
    gray = img_original.convert("L")
    inverted_gray = ImageOps.invert(gray)
    
    # Text Mask: Brighter than threshold -> Text (255)
    mask = inverted_gray.point(lambda p: 255 if p > threshold_val else 0)
    
    if blur_radius > 0:
        mask = mask.filter(ImageFilter.GaussianBlur(radius=blur_radius))
    
    mask_arr = np.array(mask).astype(float) / 255.0
    mask_arr = np.expand_dims(mask_arr, axis=2)

    bg_arr = np.full_like(np.array(img_original), theme.bg_color)
    fg_arr = np.full_like(np.array(img_original), theme.fg_color)
    
    # Composite: FG where Text, BG elsewhere
    composite_arr = (fg_arr * mask_arr + bg_arr * (1.0 - mask_arr)).astype(np.uint8)
    final_img = Image.fromarray(composite_arr)

    # 3. Smart Image Preservation
    # Detect images on the page and paste the ORIGINAL pixels back on top
    # Scale factor for DPI
    scale = dpi / 72.0
    
    # get_image_info returns a list of dicts with 'bbox'
    image_infos = page.get_image_info()
    
    for img_info in image_infos:
        bbox = img_info['bbox']
        # Convert PDF coords to Image coords
        x0, y0, x1, y1 = [int(v * scale) for v in bbox]
        
        # Clip to image bounds
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(img_original.width, x1), min(img_original.height, y1)
        
        if x1 - x0 < 5 or y1 - y0 < 5:
            continue  # Skip tiny artifacts
            
        # Crop original region
        region = img_original.crop((x0, y0, x1, y1))
        
        # Enhance for dark mode (slight dimming)
        region = enhance_image_for_dark_mode(region)
        
        # Paste back onto the dark background
        final_img.paste(region, (x0, y0))

    return final_img

def run_image_mode(doc_in, doc_out, theme, args):
    """Executes the rasterization workflow with image preservation."""
    for page in tqdm(doc_in, desc="Processing Pages (Image Mode)", unit="page"):
        processed_img = process_page_image(
            page, 
            theme, 
            args.dpi, 
            args.threshold, 
            args.blur
        )
        
        img_byte_arr = io.BytesIO()
        processed_img.save(img_byte_arr, format='JPEG', quality=85)
        
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.getvalue())

# ==============================================================================
# Vector Mode Logic (Reconstruction + Image Restoration)
# ==============================================================================

def run_vector_mode(doc_in, doc_out, theme, args):
    """Executes the vector reconstruction workflow."""
    
    # Normalize colors for PyMuPDF (0.0 to 1.0)
    bg_col_norm = (theme.bg_color[0]/255, theme.bg_color[1]/255, theme.bg_color[2]/255)
    fg_col_norm = (theme.fg_color[0]/255, theme.fg_color[1]/255, theme.fg_color[2]/255)

    for page_num, page in enumerate(tqdm(doc_in, desc="Processing Pages (Vector Mode)", unit="page")):
        
        # Create new page
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Fill Background
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_col_norm, fill=bg_col_norm)
        shape.commit()

        # 2. Restore Images
        # We use get_image_info to find where images are located
        image_list = page.get_image_info(xrefs=True)
        
        for img in image_list:
            xref = img['xref']
            bbox = img['bbox']
            
            # Skip invalid images
            if xref == 0: continue

            try:
                # Extract the image
                base_img = doc_in.extract_image(xref)
                if not base_img: continue
                
                img_bytes = base_img["image"]
                
                # We can try to optimize it (dim it) before insertion using Pillow
                # but direct insertion is faster and preserves format.
                # To match "Maximize Style", we keep it original.
                new_page.insert_image(fitz.Rect(bbox), stream=img_bytes)
            except Exception:
                pass # Skip if extraction fails

        # 3. Redraw Text with Style Preservation
        text_data = page.get_text("dict")
        
        for block in text_data.get("blocks", []):
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"]
                        size = span["size"]
                        origin = span["origin"]
                        flags = span["flags"]
                        color_int = span["color"]
                        
                        # Font mapping (Best Effort)
                        font_name = "helv" 
                        if flags & 2 ** 4: font_name = "tiro" # Serif
                        if flags & 2 ** 3: font_name = "cour" # Mono
                        
                        suffix = ""
                        if flags & 2 ** 1: suffix += "o" # Italic
                        if flags & 2 ** 2: suffix += "b" # Bold
                        full_font = font_name + suffix
                        
                        # Color Logic:
                        # If the original text is DARK (standard reading text), swap to Theme FG.
                        # If the original text is COLORED/BRIGHT (links, code highlight), preserve it.
                        span_rgb = int_to_rgb(color_int)
                        
                        final_color = fg_col_norm # Default to theme text
                        
                        if not is_dark_color(span_rgb, threshold=50):
                            # It is a colored/light element. Preserve style.
                            # Normalize 0-255 to 0-1
                            final_color = (span_rgb[0]/255, span_rgb[1]/255, span_rgb[2]/255)
                        
                        try:
                            new_page.insert_text(
                                point=origin,
                                text=text,
                                fontsize=size,
                                fontname=full_font,
                                color=final_color
                            )
                        except Exception:
                            # Fallback
                            new_page.insert_text(
                                point=origin,
                                text=text,
                                fontsize=size,
                                color=final_color
                            )

# ==============================================================================
# Main Execution
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to Dark Mode with Image Preservation.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument("input_pdf", help="Path to source PDF")
    parser.add_argument("output_pdf", help="Path to save processed PDF")
    
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled",
                        help="Color scheme selection:\n" + 
                             "\n".join([f"  {k}: {v.description}" for k, v in THEMES.items()]))
    
    parser.add_argument("--mode", choices=["image", "vector"], default="image",
                        help="Processing mode:\n" +
                             "  image:  Robust. Rasterizes pages. Preserves images perfectly.\n" +
                             "  vector: Selectable text. Reconstructs layout and inserts images.")
    
    parser.add_argument("--dpi", type=int, default=150,
                        help="Resolution for image rasterization (default: 150).")
    
    parser.add_argument("--threshold", type=int, default=100,
                        help="Luminance threshold (0-255) for text detection.")
    
    parser.add_argument("--blur", type=float, default=1.0,
                        help="Blur radius for text mask (Image Mode only).")

    args = parser.parse_args()

    try:
        doc_in = fitz.open(args.input_pdf)
    except Exception as e:
        print(f"Error opening input file: {e}")
        sys.exit(1)

    print(f"--- PDF Dark Mode Converter ---")
    print(f"Input: {args.input_pdf}")
    print(f"Theme: {THEMES[args.theme].name}")
    print(f"Mode:  {args.mode.capitalize()}")
    print("-" * 30)

    doc_out = fitz.open()

    try:
        if args.mode == "image":
            run_image_mode(doc_in, doc_out, THEMES[args.theme], args)
        else:
            run_vector_mode(doc_in, doc_out, THEMES[args.theme], args)

        doc_out.save(args.output_pdf)
        print(f"\nSuccess! Saved to: {args.output_pdf}")

    except KeyboardInterrupt:
        print("\nProcess interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\nAn error occurred during processing: {e}")
        sys.exit(1)
    finally:
        doc_in.close()
        doc_out.close()

if __name__ == "__main__":
    main()
