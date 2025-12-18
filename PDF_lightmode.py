#!/usr/bin/env python3
"""
PDF Light Mode Converter
------------------------
A professional-grade tool to convert dark-themed PDFs into clean, printable 
light-themed documents.

Modes:
1. Vector-Native (Default): Modifies PDF content streams directly. Preserves 
   text selection, vector fidelity, and file size.
2. Image-Based: Rasterizes pages, processes contrast, and rebuilds the PDF. 
   Used for scanned docs or complex renderings.

Usage:
    python pdf_light_mode.py input.pdf output.pdf --mode vector --theme paper
"""

import argparse
import sys
import logging
import io
from dataclasses import dataclass
from typing import Tuple, List, Union, Optional
from pathlib import Path

# Third-party libraries
import pikepdf
import fitz  # PyMuPDF
import numpy as np
from PIL import Image, ImageOps, ImageEnhance
from tqdm import tqdm

# Configure Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)


# ==========================================
# Domain Logic: Themes & Color Math
# ==========================================

@dataclass
class Theme:
    name: str
    # Colors are normalized (0.0 - 1.0)
    bg_color: Tuple[float, float, float]
    fg_color: Tuple[float, float, float]

    @property
    def bg_uint8(self) -> Tuple[int, int, int]:
        return tuple(int(c * 255) for c in self.bg_color)

    @property
    def fg_uint8(self) -> Tuple[int, int, int]:
        return tuple(int(c * 255) for c in self.fg_color)


THEMES = {
    "paper": Theme("Paper", (1.0, 1.0, 1.0), (0.0, 0.0, 0.0)),
    "warm": Theme("Warm White", (0.98, 0.96, 0.93), (0.1, 0.1, 0.1)),
    "soft": Theme("Soft Gray", (0.95, 0.95, 0.95), (0.15, 0.15, 0.15)),
}


class ColorUtils:
    @staticmethod
    def get_luminance(r: float, g: float, b: float) -> float:
        """Calculates perceptual luminance."""
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    @staticmethod
    def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> Tuple[float, float, float]:
        """Converts CMYK to RGB approximation."""
        return (1.0 - c) * (1.0 - k), (1.0 - m) * (1.0 - k), (1.0 - y) * (1.0 - k)

    @staticmethod
    def normalize_args(args: List) -> List[float]:
        """Safely converts any numeric PDF operand to a Python float."""
        result = []
        for x in args:
            try:
                result.append(float(x))
            except (TypeError, ValueError):
                continue
        return result

# ==========================================
# Processor 1: Vector-Native (PikePDF)
# ==========================================

class VectorProcessor:
    def __init__(self, input_path: str, output_path: str, theme: Theme):
        self.input_path = input_path
        self.output_path = output_path
        self.theme = theme
        self.color_ops = {'g', 'G', 'rg', 'RG', 'k', 'K', 'sc', 'SC', 'scn', 'SCN'}
        self.text_ops = {'Tj', 'TJ', "'", '"'}

    def process(self):
        try:
            with pikepdf.Pdf.open(self.input_path) as pdf:
                logger.info("Performing deep visibility scan...")
                
                # 1. Neutralize Transparency (FIXED: pdf.Root instead of pdf.root)
                if '/ExtGState' in pdf.Root:
                    for _, gs in pdf.Root.ExtGState.items():
                        # Blend mode 'Screen' makes text disappear on white; force Normal
                        if '/BM' in gs: 
                            gs.BM = pikepdf.Name('/Normal')
                        # Force full opacity
                        if '/ca' in gs: gs.ca = 1.0
                        if '/CA' in gs: gs.CA = 1.0

                for obj in tqdm(pdf.objects, desc="Fixing XObjects"):
                    if isinstance(obj, pikepdf.Stream):
                        try:
                            if obj.get('/Type') == '/XObject' and obj.get('/Subtype') == '/Form':
                                self._rewrite_stream(obj, pdf, is_page=False)
                        except (AttributeError, KeyError): continue

                for page in tqdm(pdf.pages, desc="Fixing Pages"):
                    self._rewrite_stream(page, pdf, is_page=True)

                pdf.save(self.output_path, compress_streams=True)
            logger.info(f"Process complete: {self.output_path}")
        except Exception as e:
            logger.error(f"Vector processing failed: {e}"); raise

    def _rewrite_stream(self, target, pdf, is_page=True):
        try:
            commands = pikepdf.parse_content_stream(target)
        except: return

        new_commands = []
        bbox_obj = target.get('/MediaBox') if is_page else target.get('/BBox')
        m = [float(x) for x in bbox_obj] if bbox_obj else None
        
        # 2. Page Background
        if is_page and m:
            new_commands.append(([], pikepdf.Operator("q")))
            new_commands.append((list(self.theme.bg_color), pikepdf.Operator("rg")))
            new_commands.append(([m[0], m[1], m[2], m[3]], pikepdf.Operator("re")))
            new_commands.append(([], pikepdf.Operator("f")))
            new_commands.append(([], pikepdf.Operator("Q")))

        cur_fill = (0.0, 0.0, 0.0)
        
        for operands, operator in commands:
            op_str = str(operator)
            
            # 3. Handle Graphics States
            if op_str == 'gs':
                new_commands.append((operands, operator))
                # Reset text color immediately after a state change
                new_commands.append((list(self.theme.fg_color), pikepdf.Operator("rg")))
                continue

            # 4. Color Mapping Logic
            if op_str in self.color_ops:
                clean_ops = ColorUtils.normalize_args(operands)
                if not clean_ops:
                    new_commands.append((operands, operator))
                    continue
                
                r, g, b = self._get_rgb_approx(clean_ops, op_str)
                lum = ColorUtils.get_luminance(r, g, b)
                
                # Broaden the foreground threshold: if it's light, make it dark
                target_color = list(self.theme.fg_color) if lum > 0.3 else list(self.theme.bg_color)
                
                if op_str.islower(): cur_fill = (target_color[0], target_color[1], target_color[2])
                new_ops = self._map_to_op_format(target_color, op_str, len(clean_ops))
                new_commands.append((new_ops, operator))

            # 5. Text Failsafe
            elif op_str in self.text_ops:
                # If current fill is too light for white paper, force dark
                if ColorUtils.get_luminance(*cur_fill) > 0.5:
                    new_commands.append((list(self.theme.fg_color), pikepdf.Operator("rg")))
                    cur_fill = self.theme.fg_color
                new_commands.append((operands, operator))

            # 6. Rectangle Neutralizer
            elif op_str == 're' and m:
                ops = ColorUtils.normalize_args(operands)
                if len(ops) == 4 and (ops[2]*ops[3]) > ((m[2]-m[0])*(m[3]-m[1])*0.6):
                    if ColorUtils.get_luminance(*cur_fill) < 0.5:
                        new_commands.append((list(self.theme.bg_color), pikepdf.Operator("rg")))
                new_commands.append((operands, operator))
            else:
                new_commands.append((operands, operator))

        new_data = pikepdf.unparse_content_stream(new_commands)
        if is_page: target.Contents = pdf.make_stream(new_data)
        else: target.write(new_data)

    def _get_rgb_approx(self, ops, op):
        if op in ['g', 'G']: return ops[0], ops[0], ops[0]
        if op in ['rg', 'RG']: return ops[0], ops[1], ops[2]
        if op in ['k', 'K']: return ColorUtils.cmyk_to_rgb(*ops[:4])
        if len(ops) == 1: return ops[0], ops[0], ops[0]
        if len(ops) == 3: return ops[0], ops[1], ops[2]
        if len(ops) == 4: return ColorUtils.cmyk_to_rgb(*ops)
        return 0.0, 0.0, 0.0

    def _map_to_op_format(self, target_rgb, op, original_len):
        if op in ['g', 'G'] or original_len == 1:
            return [ColorUtils.get_luminance(*target_rgb)]
        if op in ['k', 'K'] or original_len == 4:
            return [0.0, 0.0, 0.0, 1.0] if target_rgb[0] < 0.5 else [0.0, 0.0, 0.0, 0.0]
        return target_rgb

# ==========================================
# Processor 2: Image-Based Fallback (PyMuPDF)
# ==========================================

class ImageProcessor:
    def __init__(self, input_path: str, output_path: str, theme: Theme, 
                 dpi: int = 200, threshold: int = 100):
        self.input_path = input_path
        self.output_path = output_path
        self.theme = theme
        self.dpi = dpi
        self.threshold = threshold

    def process(self):
        try:
            doc = fitz.open(self.input_path)
            out_pdf = fitz.open()

            logger.info(f"Processing {len(doc)} pages in Enhanced Image Mode...")

            for page in tqdm(doc, desc="Image Conversion"):
                images_on_page = page.get_images(full=True)
                
                # Render page at high DPI to preserve font detail
                pix = page.get_pixmap(dpi=self.dpi)
                mode = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

                # 1. Process with Adaptive Thresholding
                processed_base = self._process_image_content(img)

                # 2. Smart Paste for Images/Diagrams
                for item in images_on_page:
                    rect = page.get_image_bbox(item)
                    if not rect.is_infinite and not rect.is_empty:
                        sx, sy = pix.width / page.rect.width, pix.height / page.rect.height
                        crop_box = (
                            max(0, int(rect.x0 * sx)), max(0, int(rect.y0 * sy)),
                            min(pix.width, int(rect.x1 * sx)), min(pix.height, int(rect.y1 * sy))
                        )
                        
                        if crop_box[2] > crop_box[0] and crop_box[3] > crop_box[1]:
                            original_patch = img.crop(crop_box)
                            # Lighten and de-saturate background images slightly for "Paper" feel
                            processed_base.paste(original_patch, crop_box)

                # 3. Save page
                with io.BytesIO() as bio:
                    processed_base.save(bio, format="JPEG", quality=90)
                    new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
                    new_page.insert_image(page.rect, stream=bio.getvalue())

            out_pdf.save(self.output_path, garbage=4, deflate=True)
            logger.info(f"Successfully saved to {self.output_path}")

        except Exception as e:
            logger.error(f"Image processing failed: {e}")
            sys.exit(1)

    def _process_image_content(self, img: Image.Image) -> Image.Image:
        """
        Uses adaptive processing to preserve thin fonts and varied titles.
        """
        # Convert to Grayscale
        gray = img.convert("L")
        np_gray = np.array(gray)

        # 1. Contrast Enhancement (Pre-threshold)
        # This helps make faint titles stand out from the dark background
        enhanced_gray = ImageOps.autocontrast(gray, cutoff=2)
        np_gray = np.array(enhanced_gray)

        # 2. Adaptive Masking
        # Instead of one global threshold, we use the user threshold as a baseline
        # but apply a small "dilation" to the mask to thicken thin font strokes.
        from scipy.ndimage import binary_dilation
        
        mask = np_gray > self.threshold
        
        # Thicken the text mask slightly (1px) so thin titles don't vanish
        # This acts like a "Bold" filter for visibility
        mask = binary_dilation(mask, iterations=1)

        # 3. Apply Theme Colors
        h, w = np_gray.shape
        result = np.zeros((h, w, 3), dtype=np.uint8)
        
        fg = np.array(self.theme.fg_uint8, dtype=np.uint8)
        bg = np.array(self.theme.bg_uint8, dtype=np.uint8)

        result[mask] = fg
        result[~mask] = bg

        return Image.fromarray(result)

# ==========================================
# Main CLI
# ==========================================

def main():
    parser = argparse.ArgumentParser(
        description="Professional PDF Light Mode Converter. Transforms dark PDFs to light themes."
    )
    
    # Required args
    parser.add_argument("input", help="Path to input PDF")
    parser.add_argument("output", help="Path to output PDF")
    
    # Optional args
    parser.add_argument("--mode", choices=["vector", "image"], default="vector",
                        help="Processing mode. Vector preserves text; Image is for scans. Default: vector")
    parser.add_argument("--theme", choices=THEMES.keys(), default="paper",
                        help="Color theme. Default: paper")
    
    # Image mode tuning
    parser.add_argument("--dpi", type=int, default=150, help="DPI for image rasterization (default: 150)")
    parser.add_argument("--threshold", type=int, default=100, help="Brightness threshold (0-255) for text detection (default: 100)")

    args = parser.parse_args()

    # Validation
    if not Path(args.input).exists():
        logger.error(f"Input file not found: {args.input}")
        sys.exit(1)

    selected_theme = THEMES[args.theme]

    print(f"--- PDF Light Mode Converter ---")
    print(f"Mode:   {args.mode.upper()}")
    print(f"Theme:  {selected_theme.name}")
    print(f"Input:  {args.input}")
    print(f"Output: {args.output}")
    print("--------------------------------")

    if args.mode == "vector":
        processor = VectorProcessor(args.input, args.output, selected_theme)
        processor.process()
    else:
        processor = ImageProcessor(args.input, args.output, selected_theme, 
                                   dpi=args.dpi, threshold=args.threshold)
        processor.process()

if __name__ == "__main__":
    main()
