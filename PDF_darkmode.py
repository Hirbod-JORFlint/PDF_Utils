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
from PIL import Image, ImageFilter, ImageOps
from tqdm import tqdm
from dataclasses import dataclass
from typing import Tuple, Optional
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
# Image Processing Logic (Raster Mode)
# ==============================================================================

def process_page_image(
    page: fitz.Page, 
    theme: Theme, 
    dpi: int, 
    threshold_val: int, 
    blur_radius: float
) -> Image.Image:
    """
    Renders a PDF page to an image and recolors it based on luminance.
    """
    # 1. Render page to pixmap
    pix = page.get_pixmap(dpi=dpi)
    img_original = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # 2. Convert to Grayscale for analysis
    # We invert logic here: usually PDF text is dark on light.
    # To detect text, we want to find dark pixels.
    # L mode: 0 is black, 255 is white.
    gray = img_original.convert("L")
    
    # 3. Create Text Mask
    # We want text to be White (255) in the mask and background to be Black (0).
    # Since standard text is dark, we invert the grayscale image first.
    inverted_gray = ImageOps.invert(gray)
    
    # Apply Threshold: Pixels brighter than X become text (255)
    # Using a point function for binary thresholding
    mask = inverted_gray.point(lambda p: 255 if p > threshold_val else 0)
    
    # 4. Soften the mask (Anti-aliasing simulation)
    if blur_radius > 0:
        mask = mask.filter(ImageFilter.GaussianBlur(radius=blur_radius))
    
    # 5. Composite Construction using Numpy for speed
    # Normalize mask to 0..1
    mask_arr = np.array(mask).astype(float) / 255.0
    mask_arr = np.expand_dims(mask_arr, axis=2) # Shape (H, W, 1)

    # Create solid color layers
    bg_arr = np.full_like(np.array(img_original), theme.bg_color)
    fg_arr = np.full_like(np.array(img_original), theme.fg_color)
    
    # Blend BG and FG based on Mask
    # Result = FG * Mask + BG * (1 - Mask)
    composite_arr = (fg_arr * mask_arr + bg_arr * (1.0 - mask_arr)).astype(np.uint8)
    composite_img = Image.fromarray(composite_arr)

    # 6. Lightly blend original image back in
    # This preserves non-text elements (images, graphs) that might have been 
    # flattened by the thresholding, albeit tinted by the theme.
    # 15% original, 85% processed
    final_img = Image.blend(composite_img, img_original, alpha=0.15)
    
    return final_img

def run_image_mode(doc_in, doc_out, theme, args):
    """Executes the rasterization workflow."""
    for page in tqdm(doc_in, desc="Processing Pages (Image Mode)", unit="page"):
        # Process visual
        processed_img = process_page_image(
            page, 
            theme, 
            args.dpi, 
            args.threshold, 
            args.blur
        )
        
        # Convert PIL image to PyMuPDF rect/stream
        img_byte_arr = io.BytesIO()
        processed_img.save(img_byte_arr, format='JPEG', quality=85)
        
        # Create new page in output doc with same dimensions
        # Note: We use the pixel dimensions of the rendered image to set page size
        # to ensure high resolution is kept, or map back to original points.
        # Here we map back to original page size.
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=img_byte_arr.getvalue())

# ==============================================================================
# Vector Processing Logic (Experimental)
# ==============================================================================

def run_vector_mode(doc_in, doc_out, theme, args):
    """Executes the experimental vector reconstruction workflow."""
    
    # Normalize colors for PyMuPDF (0.0 to 1.0)
    bg_col = (theme.bg_color[0]/255, theme.bg_color[1]/255, theme.bg_color[2]/255)
    fg_col = (theme.fg_color[0]/255, theme.fg_color[1]/255, theme.fg_color[2]/255)

    for page_num, page in enumerate(tqdm(doc_in, desc="Processing Pages (Vector Mode)", unit="page")):
        
        # Create new page
        new_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        
        # 1. Fill Background
        shape = new_page.new_shape()
        shape.draw_rect(new_page.rect)
        shape.finish(color=bg_col, fill=bg_col)
        shape.commit()

        # 2. Extract Text
        # "dict" gives detailed hierarchy: block -> line -> span
        text_data = page.get_text("dict")
        
        for block in text_data.get("blocks", []):
            if block["type"] == 0:  # Text block
                for line in block["lines"]:
                    for span in line["spans"]:
                        # Extract attributes
                        text = span["text"]
                        size = span["size"]
                        origin = span["origin"] # (x, y) baseline
                        # bbox = span["bbox"]
                        
                        # Font handling is tricky. We fallback to standard fonts.
                        # We use standard PyMuPDF fonts to avoid dependency hell.
                        font_flags = span["flags"]
                        font_name = "helv" # Default sans-serif
                        
                        if font_flags & 2 ** 0: # Superscript?
                             pass # Ignored for simplicity
                        if font_flags & 2 ** 4: # Serif
                            font_name = "tiro" # Times Roman
                        if font_flags & 2 ** 3: # Monospaced
                            font_name = "cour" # Courier
                        
                        # Add Bold/Italic suffix
                        suffix = ""
                        if font_flags & 2 ** 1: # Italic
                            suffix += "o" # 'Oblique' or Italic
                        if font_flags & 2 ** 2: # Bold
                            suffix += "b"
                        
                        full_font = font_name + suffix
                        
                        # 3. Redraw Text
                        try:
                            new_page.insert_text(
                                point=origin,
                                text=text,
                                fontsize=size,
                                fontname=full_font,
                                color=fg_col
                            )
                        except Exception:
                            # Fallback if font combo fails
                            new_page.insert_text(
                                point=origin,
                                text=text,
                                fontsize=size,
                                color=fg_col
                            )

        # Note: We deliberately do not re-draw images in vector mode 
        # as per "document that images are not re-drawn".
        
    print("\n[!] Note: Vector mode does not preserve images, vector graphics, or complex formatting tables.")

# ==============================================================================
# Main Execution
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to Dark Mode.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument("input_pdf", help="Path to source PDF")
    parser.add_argument("output_pdf", help="Path to save processed PDF")
    
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled",
                        help="Color scheme selection:\n" + 
                             "\n".join([f"  {k}: {v.description}" for k, v in THEMES.items()]))
    
    parser.add_argument("--mode", choices=["image", "vector"], default="image",
                        help="Processing mode:\n" +
                             "  image:  (Default) Robust. Rasterizes pages. Non-selectable text.\n" +
                             "  vector: Experimental. Re-types text. Selectable text, loses images.")
    
    parser.add_argument("--dpi", type=int, default=150,
                        help="Resolution for image rasterization (default: 150). Higher = slower but sharper.")
    
    parser.add_argument("--threshold", type=int, default=100,
                        help="Luminance threshold (0-255) to detect text pixels (default: 100).")
    
    parser.add_argument("--blur", type=float, default=1.0,
                        help="Gaussian blur radius for text mask (default: 1.0). Softens jagged edges.")

    args = parser.parse_args()

    # File Checks
    try:
        doc_in = fitz.open(args.input_pdf)
    except Exception as e:
        print(f"Error opening input file: {e}")
        sys.exit(1)

    print(f"--- PDF Dark Mode Converter ---")
    print(f"Input: {args.input_pdf}")
    print(f"Theme: {THEMES[args.theme].name}")
    print(f"Mode:  {args.mode.capitalize()}")
    if args.mode == "image":
        print(f"DPI:   {args.dpi}")
    print("-" * 30)

    # Initialize Output Document
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
