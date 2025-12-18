#!/usr/bin/env python3
"""
PDF Dark Mode Converter (v8.0 - Native Content Editing)
=======================================================
Optimized for perfect font preservation. Instead of rewriting text,
this version edits the PDF content streams directly using PikePDF.
This ensures Bold, Italic, Symbols, and Layouts remain 100% original.

Dependencies:
    pip install pikepdf pymupdf pillow numpy tqdm
"""

import argparse
import sys
import io
import fitz  # PyMuPDF
import pikepdf
from pikepdf import Pdf, Name, Operator, Stream
import numpy as np
from PIL import Image, ImageFilter, ImageOps, ImageEnhance
from tqdm import tqdm
from dataclasses import dataclass
from typing import Tuple, List, Union

@dataclass
class Theme:
    name: str
    bg_color: Tuple[float, float, float] # Normalized 0.0-1.0
    fg_color: Tuple[float, float, float] # Normalized 0.0-1.0

# Define Themes (Normalized RGB)
THEMES = {
    'amoled': Theme('AMOLED', (0.0, 0.0, 0.0), (0.92, 0.92, 0.92)),
    'sepia':  Theme('Sepia',  (0.17, 0.12, 0.12), (0.86, 0.80, 0.74)),
    'navy':   Theme('Navy',   (0.04, 0.10, 0.18), (0.39, 1.0, 0.85))
}

def is_dark_color(operands: List) -> bool:
    """Detects if a color operation is 'dark' (likely text)."""
    # Operands are lists of numbers. Logic depends on color space (Gray, RGB, CMYK)
    if not operands: return False
    
    try:
        nums = [float(x) for x in operands]
        if len(nums) == 1:   # Grayscale (0=Black, 1=White)
            return nums[0] < 0.5
        elif len(nums) == 3: # RGB
            return (0.299*nums[0] + 0.587*nums[1] + 0.114*nums[2]) < 0.5
        elif len(nums) == 4: # CMYK (0=White, 1=Black usually, but K dominates)
            # Simple heuristic: if K is high or CMY is high, it's dark
            return nums[3] > 0.5 or (nums[0]+nums[1]+nums[2]) > 1.5
    except:
        return True # Assume text if unclear
    return False

def create_solid_background(page_rect, color_rgb):
    """Creates a background rectangle command stream."""
    x, y, w, h = page_rect
    r, g, b = color_rgb
    
    # Pikepdf expects content stream instructions as (operands, operator) tuples
    # Operands must be a list, even if empty.
    return [
        ([], Operator("q")),               # Save graphics state
        ([r, g, b], Operator("rg")),       # Set non-stroking color (RGB)
        ([x, y, w, h], Operator("re")),    # Rectangle
        ([], Operator("f")),               # Fill path
        ([], Operator("Q"))                # Restore graphics state
    ]

def run_vector_mode_native(input_path, output_path, theme_key):
    """
    Uses PikePDF to edit content streams in-place.
    """
    theme = THEMES[theme_key]
    pdf = Pdf.open(input_path)
    
    # Operators that define color
    COLOR_OPS = {'g', 'G', 'rg', 'RG', 'k', 'K'}

    for page in tqdm(pdf.pages, desc="Processing (Vector Native)", unit="page"):
        # 1. Parse existing content stream
        # This coalesces multiple content streams into one list of commands
        commands = pikepdf.parse_content_stream(page)
        new_commands = []
        
        # 2. Add Background (Prepend to stream)
        mediabox = [float(c) for c in page.MediaBox]
        bg_cmds = create_solid_background(mediabox, theme.bg_color)
        new_commands.extend(bg_cmds)

        # 3. Process existing commands
        for operands, operator in commands:
            op_name = str(operator)
            
            if op_name in COLOR_OPS:
                if is_dark_color(operands):
                    # Change dark text/lines to Theme Foreground
                    new_op = 'RG' if op_name in ['G', 'RG', 'K'] else 'rg'
                    new_commands.append((list(theme.fg_color), Operator(new_op)))
                else:
                    # Change light background artifacts to Theme Background
                    new_op = 'RG' if op_name in ['G', 'RG', 'K'] else 'rg'
                    new_commands.append((list(theme.bg_color), Operator(new_op)))
            else:
                # Keep text positioning, fonts, and images exactly as is
                new_commands.append((operands, operator))

        # 4. Write back the modified stream
        # FIX: Use pdf.make_stream() and assign to page.Contents
        new_content_bytes = pikepdf.unparse_content_stream(new_commands)
        page.Contents = pdf.make_stream(new_content_bytes)

    # Save
    pdf.save(output_path)


def run_image_mode(doc_in, doc_out, theme_key, args):
    """Fallback Image-based processing using PyMuPDF + Pillow."""
    # Convert normalized theme float to 0-255 int
    theme = THEMES[theme_key]
    theme_bg = np.array([int(c*255) for c in theme.bg_color], dtype=np.uint8)
    theme_fg = np.array([int(c*255) for c in theme.fg_color], dtype=np.uint8)

    for page in tqdm(doc_in, desc="Processing (Image Mode)", unit="page"):
        pix = page.get_pixmap(dpi=args.dpi)
        img_orig = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Create mask
        gray = ImageOps.invert(img_orig.convert("L"))
        mask = gray.point(lambda p: 255 if p > args.threshold else 0)
        if args.blur > 0: mask = mask.filter(ImageFilter.GaussianBlur(args.blur))
        
        # Composite
        m_arr = np.expand_dims(np.array(mask).astype(float) / 255.0, axis=2)
        canvas_bg = np.full_like(np.array(img_orig), theme_bg)
        canvas_fg = np.full_like(np.array(img_orig), theme_fg)
        processed_img = Image.fromarray((canvas_fg * m_arr + canvas_bg * (1.0 - m_arr)).astype(np.uint8))

        # Image preservation (simple crop and paste back)
        scale = args.dpi / 72.0
        for info in page.get_image_info():
            b = [int(v * scale) for v in info['bbox']]
            if b[2] > b[0] and b[3] > b[1]:
                crop = img_orig.crop((max(0, b[0]), max(0, b[1]), min(img_orig.width, b[2]), min(img_orig.height, b[3])))
                # Slightly dim images to blend better
                processed_img.paste(ImageEnhance.Brightness(crop).enhance(0.85), (b[0], b[1]))

        out_page = doc_out.new_page(width=page.rect.width, height=page.rect.height)
        buf = io.BytesIO()
        processed_img.save(buf, format="JPEG", quality=85)
        out_page.insert_image(out_page.rect, stream=buf.getvalue())

def main():
    parser = argparse.ArgumentParser(description="PDF Dark Mode Converter v8.0")
    parser.add_argument("input", help="Input PDF path")
    parser.add_argument("output", help="Output PDF path")
    parser.add_argument("--theme", choices=THEMES.keys(), default="amoled")
    parser.add_argument("--mode", choices=["image", "vector"], default="vector")
    
    # Image mode specific args
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--threshold", type=int, default=128)
    parser.add_argument("--blur", type=float, default=0.5)

    args = parser.parse_args()
    
    try:
        if args.mode == "vector":
            # Use PikePDF for Vector Mode (Best Preservation)
            run_vector_mode_native(args.input, args.output, args.theme)
        else:
            # Use PyMuPDF for Image Mode (Best for Scans/Complex Layouts)
            doc_in, doc_out = fitz.open(args.input), fitz.open()
            run_image_mode(doc_in, doc_out, args.theme, args)
            doc_out.save(args.output, deflate=True, garbage=4)
            doc_in.close()
            doc_out.close()
            
        print(f"\n[+] Success! File saved to: {args.output}")
    except Exception as e:
        print(f"\n[!] Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
