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
    """Helper methods for PDF color space conversions and luminance logic."""

    @staticmethod
    def get_luminance(r: float, g: float, b: float) -> float:
        """Calculates perceptual luminance (Rec. 709)."""
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    @staticmethod
    def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> Tuple[float, float, float]:
        """Simple CMYK to RGB conversion."""
        r = (1.0 - c) * (1.0 - k)
        g = (1.0 - m) * (1.0 - k)
        b = (1.0 - y) * (1.0 - k)
        return r, g, b

    @staticmethod
    def normalize_args(args: List) -> List[float]:
        """
        Converts PikePDF arguments to standard python floats.
        Handles integers, floats, and pikepdf numeric objects.
        """
        return [float(x) for x in args]


# ==========================================
# Processor 1: Vector-Native (PikePDF)
# ==========================================

class VectorProcessor:
    def __init__(self, input_path: str, output_path: str, theme: Theme):
        self.input_path = input_path
        self.output_path = output_path
        self.theme = theme
        self.color_ops = {'g', 'G', 'rg', 'RG', 'k', 'K'}

    def process(self):
        try:
            with pikepdf.Pdf.open(self.input_path) as pdf:
                logger.info("Scanning for XObjects and Page streams...")
                
                # 1. Iterate over all indirect objects to find Form XObjects
                # Fix: pdf.objects is a sequence, not a dict
                for obj in tqdm(pdf.objects, desc="Processing XObjects"):
                    if isinstance(obj, pikepdf.Stream):
                        # Safely check types without triggering exceptions on non-dict streams
                        try:
                            if obj.get('/Type') == '/XObject' and obj.get('/Subtype') == '/Form':
                                self._rewrite_stream(obj, pdf, is_page=False)
                        except (AttributeError, KeyError):
                            continue

                # 2. Process Page Streams
                for page in tqdm(pdf.pages, desc="Processing Pages"):
                    self._rewrite_stream(page, pdf, is_page=True)

                # Save with compression and object stream generation
                pdf.save(self.output_path, compress_streams=True, 
                         object_stream_mode=pikepdf.ObjectStreamMode.generate)
                
            logger.info(f"Successfully saved to {self.output_path}")
        except Exception as e:
            logger.error(f"Vector processing failed: {e}")
            raise

    def _rewrite_stream(self, target, pdf, is_page=True):
        """
        Universal stream rewriter.
        target: pikepdf.Page (if is_page) or pikepdf.Stream (if XObject)
        """
        try:
            # parse_content_stream works on both Page and Stream objects
            commands = pikepdf.parse_content_stream(target)
        except Exception:
            return 

        new_commands = []
        
        # Determine boundaries for Background Neutralization
        # Pages use MediaBox; XObjects use BBox
        bbox_obj = target.get('/MediaBox') if is_page else target.get('/BBox')
        
        if bbox_obj:
            m = [float(x) for x in bbox_obj]
            w, h = m[2] - m[0], m[3] - m[1]
            area_threshold = (w * h) * 0.8
        else:
            m, area_threshold = None, 0

        # Prepend a full-page theme background for Page objects
        if is_page and m:
            new_commands.append(([], pikepdf.Operator("q")))
            new_commands.append((list(self.theme.bg_color), pikepdf.Operator("rg")))
            new_commands.append(([m[0], m[1], m[2], m[3]], pikepdf.Operator("re")))
            new_commands.append(([], pikepdf.Operator("f")))
            new_commands.append(([], pikepdf.Operator("Q")))

        current_fill_color = (0.0, 0.0, 0.0)

        for operands, operator in commands:
            op_str = str(operator)
            
            if op_str in self.color_ops:
                clean_ops = ColorUtils.normalize_args(operands)
                transformed = self._transform_color(clean_ops, op_str)
                # Keep track of fill color for rectangle neutralization
                if op_str in ['g', 'rg', 'k']:
                    current_fill_color = self._get_rgb_approx(clean_ops, op_str)
                new_commands.append((transformed, operator))
            
            # Neutralize large dark rectangles
            elif op_str == 're' and area_threshold > 0:
                ops = ColorUtils.normalize_args(operands)
                if (ops[2] * ops[3]) > area_threshold:
                    lum = ColorUtils.get_luminance(*current_fill_color)
                    if lum < 0.5: # If it's a large dark background shape
                        new_commands.append((list(self.theme.bg_color), pikepdf.Operator("rg")))
                new_commands.append((operands, operator))
            else:
                new_commands.append((operands, operator))

        # Write data back to the correct location
        new_data = pikepdf.unparse_content_stream(new_commands)
        if is_page:
            target.Contents = pdf.make_stream(new_data)
        else:
            # For XObjects, we write directly to the stream object
            target.write(new_data)

    def _get_rgb_approx(self, ops, op):
        """Internal helper to track the current fill color for the Neutralizer."""
        if op == 'g': return (ops[0], ops[0], ops[0])
        if op == 'rg': return (ops[0], ops[1], ops[2])
        if op == 'k': return ColorUtils.cmyk_to_rgb(*ops)
        return (0, 0, 0)

    def _transform_color(self, operands, operator):
        # Maps color to theme based on luminance
        r, g, b = self._get_rgb_approx(operands, operator)
        lum = ColorUtils.get_luminance(r, g, b)
        
        target = list(self.theme.fg_color) if lum > 0.5 else list(self.theme.bg_color)

        if operator in ['g', 'G']: return [ColorUtils.get_luminance(*target)]
        if operator in ['rg', 'RG']: return target
        if operator in ['k', 'K']: 
            # Simplified CMYK mapping
            return [0, 0, 0, 1] if lum > 0.5 else [0, 0, 0, 0]
        return operands


# ==========================================
# Processor 2: Image-Based Fallback (PyMuPDF)
# ==========================================

class ImageProcessor:
    """
    Rasterizes pages, performs computer vision operations to swap themes,
    and reconstructs the PDF. Handles scanned docs or complex graphics.
    """

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

            logger.info(f"Processing {len(doc)} pages in Image Mode (DPI={self.dpi})...")

            for page in tqdm(doc, desc="Image Conversion"):
                # 1. Detect Images (Photos/Diagrams) to preserve
                # We do this before rendering to get coordinates
                images_on_page = page.get_images(full=True)
                image_rects = []
                
                # Render full page
                pix = page.get_pixmap(dpi=self.dpi)
                mode = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)

                # 2. Process the base image (The "Text" layer)
                processed_base = self._process_image_content(img)

                # 3. Paste original images back (Heuristic preservation)
                # Note: Extracting exact bbox from get_images is tricky due to transforms.
                # A safer bet for generic PDFs is processing the whole image, 
                # but "preserving" diagrams usually requires knowing where they are.
                # We will perform a "smart paste" by detecting non-text blocks if requested,
                # but for robustness, we simply boost the contrast of the whole result.
                
                # However, to strictly follow instructions: "Detect embedded images... Crop... Paste back"
                # We use PyMuPDF's image list to locate them roughly. 
                # Since mapping PDF coordinates to the rendered Pixmap is complex,
                # we will rely on PyMuPDF's `page.get_image_bbox`
                
                for item in images_on_page:
                    xref = item[0]
                    rect = page.get_image_bbox(item)
                    
                    if not rect.is_infinite and not rect.is_empty:
                        # Convert PDF rect to Pixmap coordinates
                        # PDF user units -> Pixmap pixels
                        # Scale factor
                        sx = pix.width / page.rect.width
                        sy = pix.height / page.rect.height
                        
                        crop_box = (
                            int(rect.x0 * sx), int(rect.y0 * sy),
                            int(rect.x1 * sx), int(rect.y1 * sy)
                        )
                        
                        # Ensure crop is within bounds
                        crop_box = (
                            max(0, crop_box[0]), max(0, crop_box[1]),
                            min(pix.width, crop_box[2]), min(pix.height, crop_box[3])
                        )
                        
                        if crop_box[2] > crop_box[0] and crop_box[3] > crop_box[1]:
                            # Crop original region
                            original_patch = img.crop(crop_box)
                            
                            # Enhance the patch (slightly increase brightness/contrast)
                            enhancer = ImageEnhance.Brightness(original_patch)
                            original_patch = enhancer.enhance(1.1)
                            enhancer = ImageEnhance.Contrast(original_patch)
                            original_patch = enhancer.enhance(1.1)
                            
                            # Paste back onto processed base
                            processed_base.paste(original_patch, crop_box)

                # 4. Add page to output PDF
                with io.BytesIO() as bio:
                    processed_base.save(bio, format="JPEG", quality=85)
                    img_bytes = bio.getvalue()
                    
                    new_page = out_pdf.new_page(width=page.rect.width, height=page.rect.height)
                    new_page.insert_image(page.rect, stream=img_bytes)

            out_pdf.save(self.output_path, garbage=4, deflate=True)
            logger.info(f"Successfully saved to {self.output_path}")

        except Exception as e:
            logger.error(f"Image processing failed: {e}")
            sys.exit(1)

    def _process_image_content(self, img: Image.Image) -> Image.Image:
        """
        Thresholds the image to separate text from bg.
        Reconstructs using Theme colors.
        """
        # Convert to Grayscale
        gray = img.convert("L")
        np_gray = np.array(gray)

        # Thresholding: 
        # In Dark Mode PDF: Text is Bright (high val), BG is Dark (low val).
        # Mask = pixels > threshold (The Text)
        mask = np_gray > self.threshold

        # Create output array
        # Shape (H, W, 3)
        h, w = np_gray.shape
        result = np.zeros((h, w, 3), dtype=np.uint8)

        # Apply Theme Colors
        # Where mask is True (Text), use Foreground
        # Where mask is False (Background), use Background
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
