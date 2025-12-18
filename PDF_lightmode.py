#!/usr/bin/env python3
"""
PDF Light Mode Converter

A professional-grade tool to convert dark-themed PDFs to light-themed PDFs
while preserving vector text, fonts, layout, and vector fidelity.

Modes
-----
- Vector-Native Mode (default): Uses pikepdf to operate directly on page
  content streams, modifying color operators (g, G, rg, RG, k, K) and
  prepending a page-sized light rectangle. This mode avoids rasterization
  and preserves fonts, glyphs, positioning and vector geometry.

- Image-Based Fallback Mode: For scanned/raster-heavy PDFs this mode uses
  PyMuPDF (fitz) + Pillow + NumPy to rasterize pages, detect high-contrast
  light-on-dark text, and composite a light-theme result. Detected image
  regions are brightened slightly and pasted back to preserve photos/diagrams.

Design notes / constraints
--------------------------
- The converter operates at the content-stream level in vector mode and DOES
  NOT attempt OCR, reflow, or font substitution. It avoids re-inserting text.
- The implementation focuses on readability and correctness rather than raw
  performance; large PDFs may be slow.

Dependencies
------------
- pikepdf (vector mode)
- PyMuPDF (fitz) (image fallback mode)
- Pillow
- NumPy
- tqdm

Install via: pip install pikepdf pymupdf Pillow numpy tqdm

"""

from __future__ import annotations

import argparse
import io
import math
import re
import sys
import logging
from dataclasses import dataclass
from typing import Tuple, Callable

from tqdm import tqdm

# We'll import heavy optional dependencies lazily inside the functions that need them

# ----------------------------- Theme dataclass -----------------------------

@dataclass
class Theme:
    name: str
    bg_color: Tuple[float, float, float]  # normalized RGB (0.0 - 1.0)
    fg_color: Tuple[float, float, float]


# Preset themes
THEMES = {
    "paper": Theme("Paper", (1.0, 1.0, 1.0), (0.0, 0.0, 0.0)),
    "warm_white": Theme("Warm White", (0.992, 0.973, 0.925), (0.06, 0.06, 0.06)),
    "soft_gray": Theme("Soft Gray", (0.96, 0.96, 0.97), (0.08, 0.08, 0.08)),
}


# ----------------------------- Utilities ----------------------------------

def clamp01(v: float) -> float:
    return max(0.0, min(1.0, v))


def rgb_luminance(r: float, g: float, b: float) -> float:
    """Compute perceptual luminance in linear 0..1 space.

    The content streams usually contain sRGB values. For simplicity we use
    a standard approximate perceptual weighting.
    """
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> Tuple[float, float, float]:
    # Acrobat/most PDFs: value range 0..1
    r = 1.0 - min(1.0, c + k)
    g = 1.0 - min(1.0, m + k)
    b = 1.0 - min(1.0, y + k)
    return (r, g, b)


def format_rgb_triplet(rgb: Tuple[float, float, float]) -> str:
    # PDF content streams usually keep numbers short; 6 decimal places is safe.
    return "{:.6f} {:.6f} {:.6f} rg".format(*rgb)


def format_rgb_triplet_for_stroke(rgb: Tuple[float, float, float]) -> str:
    return "{:.6f} {:.6f} {:.6f} RG".format(*rgb)


def format_gray(val: float) -> str:
    return "{:.6f} g".format(val)


def format_gray_stroke(val: float) -> str:
    return "{:.6f} G".format(val)


def format_cmyk(cmyk: Tuple[float, float, float, float], stroke: bool) -> str:
    if stroke:
        return "{:.6f} {:.6f} {:.6f} {:.6f} K".format(*cmyk)
    else:
        return "{:.6f} {:.6f} {:.6f} {:.6f} k".format(*cmyk)


# Regex patterns to find color operators with numeric operands.
# We keep the regex conservative: numbers followed by whitespace then operator.
NUM = r"[+-]?(?:\d*\.\d+|\d+)(?:[eE][+-]?\d+)?"
RE_RGB = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+rg\b")
RE_RGB_STROKE = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+RG\b")
RE_GRAY = re.compile(rf"({NUM})\s+g\b")
RE_GRAY_STROKE = re.compile(rf"({NUM})\s+G\b")
RE_CMYK = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+k\b")
RE_CMYK_STROKE = re.compile(rf"({NUM})\s+({NUM})\s+({NUM})\s+({NUM})\s+K\b")


# ----------------------------- Vector Mode --------------------------------

def vector_mode_transform(pdf_path: str, out_path: str, theme: Theme) -> None:
    """
    Vector-native transform using pikepdf.

    Important: this function only edits content streams at the token level and
    does NOT perform rasterization, OCR, font reconstruction, or text reflow.

    The algorithm:
    - For each page read the content stream bytes (attempts multiple read
      strategies for compatibility with different pikepdf versions)
    - Prepend a small content block that paints a page-sized rectangle with the
      theme background color (wrapped in graphics state save/restore q/Q)
    - Replace explicit RGB/gray/CMYK color-setting operators (rg/RG/g/G/k/K)
      by mapping "light" colors -> theme.fg_color and "dark" colors -> theme.bg_color
    - Write content back to the page

    NOTE: This is heuristic-based. Some PDFs use colors through ColorSpace
    objects, extended graphic states, or separation names; those cases may
    require additional handling.
    """
    try:
        import pikepdf
    except Exception as e:
        logging.error("pikepdf is required for vector mode: %s", e)
        raise

    pdf = pikepdf.Pdf.open(pdf_path, allow_overwriting_input=False)

    # We will process pages and mutate content streams.
    for pnum, page in enumerate(tqdm(list(pdf.pages), desc="Vector pages")):
        # Read content bytes robustly
        raw_bytes = b""
        try:
            # Try PageContents helper if available
            try:
                pc = pikepdf.PageContents(page)
                # Different pikepdf versions expose different methods
                if hasattr(pc, "read_bytes"):
                    raw_bytes = pc.read_bytes()
                elif hasattr(pc, "streams"):
                    # streams is a list of Stream objects
                    raw_bytes = b"".join(s.read_bytes() for s in pc.streams)
                else:
                    # fallback: try iterate
                    raw_bytes = b"".join([s.read_bytes() for s in pc])
            except Exception:
                # Fallback: access page.obj["/Contents"] directly
                contents = page.obj.get("/Contents")
                if contents is None:
                    raw_bytes = b""
                elif isinstance(contents, pikepdf.Stream):
                    raw_bytes = contents.read_bytes()
                else:
                    # Possibly an array
                    if hasattr(contents, "items"):
                        raw_bytes = b"".join([c.read_bytes() for c in contents])
                    else:
                        raw_bytes = b"".join([c.read_bytes() for c in contents])
        except Exception as e:
            logging.warning("Failed to read content stream for page %d: %s", pnum + 1, e)
            raw_bytes = b""

        # Decode as latin-1 to preserve bytes 0-255 as-is in Python str.
        cs_text = raw_bytes.decode("latin-1")

        # Prepend background rectangle at the beginning of the content stream.
        # We'll get the page dimensions from MediaBox or CropBox (in user space units = points)
        try:
            mediabox = page.MediaBox
            # mediabox is an array-like: [llx, lly, urx, ury]
            llx, lly, urx, ury = [float(m) for m in mediabox]
            width = urx - llx
            height = ury - lly
        except Exception:
            # Fallback to standard A4 size if something is malformed
            width, height = 595.0, 842.0

        bg_r, bg_g, bg_b = theme.bg_color
        # Build rectangle paint sequence. We wrap in q/Q to preserve graphic state.
        # Explanation of operators used:
        # - q/Q : save/restore graphics state
        # - rg : set fill color in sRGB
        # - re : append rectangle to current path
        # - f  : fill path
        # We place this at the beginning so subsequent page content overlays on top.
        bg_cmds = []
        bg_cmds.append("q")
        bg_cmds.append("{:.6f} {:.6f} {:.6f} rg".format(bg_r, bg_g, bg_b))
        # The rectangle command: x y w h re
        bg_cmds.append(f"0 0 {width:.4f} {height:.4f} re")
        bg_cmds.append("f")
        bg_cmds.append("Q")
        bg_block = "\n".join(bg_cmds) + "\n"

        # Apply regex replacements for color operators.
        # Replacement function that decides mapping based on luminance.
        def rgb_replacer(m: re.Match, stroke: bool = False) -> str:
            r = float(m.group(1))
            g = float(m.group(2))
            b = float(m.group(3))
            # clamp in case of weird numbers
            r, g, b = clamp01(r), clamp01(g), clamp01(b)
            lum = rgb_luminance(r, g, b)
            # Heuristic threshold: > 0.6 considered light (e.g., white-ish)
            threshold = 0.6
            if lum > threshold:
                # light content => likely foreground text on dark background -> map to theme.fg_color
                target_rgb = theme.fg_color
            else:
                # dark content => likely background or dark shapes -> map to theme.bg_color
                target_rgb = theme.bg_color
            if stroke:
                return format_rgb_triplet_for_stroke(target_rgb)
            else:
                return format_rgb_triplet(target_rgb)

        def gray_replacer(m: re.Match, stroke: bool = False) -> str:
            v = float(m.group(1))
            v = clamp01(v)
            # In PDF gray 0 = black, 1 = white
            lum = v
            threshold = 0.6
            if lum > threshold:
                target_rgb = theme.fg_color
            else:
                target_rgb = theme.bg_color
            # prefer emitting rgb operator so we can preserve richer color choice
            if stroke:
                return format_rgb_triplet_for_stroke(target_rgb)
            else:
                return format_rgb_triplet(target_rgb)

        def cmyk_replacer(m: re.Match, stroke: bool = False) -> str:
            c = clamp01(float(m.group(1)))
            m_ = clamp01(float(m.group(2)))
            y = clamp01(float(m.group(3)))
            k = clamp01(float(m.group(4)))
            r, g, b = cmyk_to_rgb(c, m_, y, k)
            lum = rgb_luminance(r, g, b)
            threshold = 0.6
            if lum > threshold:
                target_rgb = theme.fg_color
            else:
                target_rgb = theme.bg_color
            # Emit rgb operator to avoid dealing with CMYK color spaces complexity
            if stroke:
                return format_rgb_triplet_for_stroke(target_rgb)
            else:
                return format_rgb_triplet(target_rgb)

        # Perform substitutions
        # Order matters to avoid overlapping replacements.
        cs_text = RE_RGB_STROKE.sub(lambda m: rgb_replacer(m, stroke=True), cs_text)
        cs_text = RE_RGB.sub(lambda m: rgb_replacer(m, stroke=False), cs_text)
        cs_text = RE_GRAY_STROKE.sub(lambda m: gray_replacer(m, stroke=True), cs_text)
        cs_text = RE_GRAY.sub(lambda m: gray_replacer(m, stroke=False), cs_text)
        cs_text = RE_CMYK_STROKE.sub(lambda m: cmyk_replacer(m, stroke=True), cs_text)
        cs_text = RE_CMYK.sub(lambda m: cmyk_replacer(m, stroke=False), cs_text)

        # Compose final content stream
        new_cs_text = bg_block + cs_text
        new_bytes = new_cs_text.encode("latin-1")

        # Write bytes back to page content.
        try:
            # Try PageContents write helpers
            try:
                pc = pikepdf.PageContents(page)
                if hasattr(pc, "write_bytes"):
                    pc.write_bytes(new_bytes)
                else:
                    # Some versions may not support write_bytes; overwrite via /Contents
                    page.obj["/Contents"] = pdf.make_stream(new_bytes)
            except Exception:
                page.obj["/Contents"] = pdf.make_stream(new_bytes)
        except Exception as e:
            logging.error("Failed to write modified content stream to page %d: %s", pnum + 1, e)

    # Save PDF with garbage collection and compression when supported
    try:
        # pikepdf supports some save options; try the most common ones and fall back
        pdf.save(out_path, optimize_streams=True, garbage_collect=True)
    except TypeError:
        try:
            pdf.save(out_path, optimize_streams=True)
        except Exception:
            pdf.save(out_path)
    except Exception:
        pdf.save(out_path)

    pdf.close()


# ----------------------------- Image Mode ---------------------------------

def image_mode_transform(pdf_path: str, out_path: str, theme: Theme, dpi: int = 150, threshold: int = 180, blur_radius: float = 0.0) -> None:
    """
    Image-based fallback transformation for raster-heavy pages.

    Steps:
    - Use PyMuPDF to rasterize each page at the requested DPI
    - Convert to grayscale and detect bright text on dark backgrounds by thresholding
    - Optionally blur the mask to soften edges
    - Composite a new image with theme background and foreground colors
    - Detect large image-like regions (high local variance) and paste enhanced
      versions of the original into the composite to preserve photographs
    - Recreate a new PDF composed of full-page images using PyMuPDF

    Notes:
    - This mode rasterizes pages; vector fidelity is not preserved but the
      result is optimized for readability.
    - No OCR or font reconstruction is performed.
    """
    try:
        import fitz  # PyMuPDF
        from PIL import Image, ImageFilter, ImageEnhance
        import numpy as np
    except Exception as e:
        logging.error("PyMuPDF, Pillow and NumPy are required for image mode: %s", e)
        raise

    src = fitz.open(pdf_path)
    out = fitz.open()

    for pnum in range(src.page_count):
        page = src.load_page(pnum)
        # Render the page at the requested DPI
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # Convert to grayscale and compute mask of bright pixels
        gray = img.convert("L")
        arr = np.array(gray).astype(np.uint8)

        # Bright text on dark backgrounds will have high pixel values where text is.
        # We create a mask for pixels brighter than the threshold.
        mask = arr > threshold

        if blur_radius and blur_radius > 0.0:
            mask_img = Image.fromarray((mask * 255).astype(np.uint8))
            mask_img = mask_img.filter(ImageFilter.GaussianBlur(radius=blur_radius))
            mask = np.array(mask_img) > 128

        # Create base composite with background color
        w, h = img.size
        bg_rgb = tuple(int(255 * clamp01(c)) for c in theme.bg_color)
        fg_rgb = tuple(int(255 * clamp01(c)) for c in theme.fg_color)

        base = Image.new("RGB", (w, h), color=bg_rgb)
        base_arr = np.array(base)

        # Where mask is True (bright text), paint foreground color
        base_arr[mask] = fg_rgb

        composite = Image.fromarray(base_arr.astype(np.uint8))

        # Attempt to detect photographic regions to preserve them.
        # Heuristic: compute local variance; regions with high variance and a large
        # bounding box are likely photographs.
        # We'll use a simple sliding-window approach with downsampling for speed.
        gray_f = arr.astype(np.float32)
        # Downsample for variance estimation
        ds = 8
        small = gray_f.reshape((h // ds, ds, w // ds, ds)).mean(axis=(1, 3))
        var_map = (gray_f.reshape((h // ds, ds, w // ds, ds)).var(axis=(1, 3)))
        # Regions with variance above threshold_var are candidates
        threshold_var = 400.0
        photomask_small = var_map > threshold_var

        # Upscale photomask
        photomask = np.kron(photomask_small, np.ones((ds, ds), dtype=bool))
        photomask = photomask[:h, :w]

        # Find bounding boxes of connected components in photomask where a large
        # fraction of pixels are not dominated by text mask. We simply iterate
        # over coarse blocks to keep memory/time reasonable.
        from scipy import ndimage  # SciPy is optional but common for this step

        try:
            labeled, ncomp = ndimage.label(photomask)
            objects = ndimage.find_objects(labeled)
        except Exception:
            # If SciPy isn't available, skip sophisticated detection and assume no images
            objects = []

        # Enhance and paste back original photographic areas (if any)
        for obj in objects:
            if obj is None:
                continue
            y0, y1 = obj[0].start, obj[0].stop
            x0, x1 = obj[1].start, obj[1].stop
            wbox = x1 - x0
            hbox = y1 - y0
            if wbox < 50 or hbox < 50:
                continue
            # Crop original region, enhance brightness/contrast slightly
            region = img.crop((x0, y0, x1, y1))
            enhancer = ImageEnhance.Brightness(region)
            region = enhancer.enhance(1.05)
            enhancer = ImageEnhance.Contrast(region)
            region = enhancer.enhance(1.08)
            composite.paste(region, (x0, y0))

        # Save composite to bytes and insert as a full page in the output PDF
        bio = io.BytesIO()
        composite.save(bio, format="PNG")
        img_bytes = bio.getvalue()

        # Create a new page with the same size in points (72 DPI units)
        page_width_pt = page.rect.width
        page_height_pt = page.rect.height
        new_page = out.new_page(width=page_width_pt, height=page_height_pt)
        # Insert image to exactly cover the page. PyMuPDF's insert_image accepts stream.
        rect = fitz.Rect(0, 0, page_width_pt, page_height_pt)
        new_page.insert_image(rect, stream=img_bytes)

    # Save output PDF
    out.save(out_path, garbage=4)
    out.close()
    src.close()


# ----------------------------- CLI ----------------------------------------


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PDF Light Mode Converter")
    parser.add_argument("input", help="Input PDF file path")
    parser.add_argument("output", help="Output PDF file path")
    parser.add_argument("--mode", choices=("vector", "image"), default="vector",
                        help="Processing mode: 'vector' (default) or 'image' (fallback)")
    parser.add_argument("--theme", choices=list(THEMES.keys()), default="paper",
                        help="Theme name to use for conversion")
    parser.add_argument("--dpi", type=int, default=150,
                        help="DPI for image mode rasterization")
    parser.add_argument("--threshold", type=int, default=180,
                        help="Grayscale threshold (0-255) to detect bright text in image mode")
    parser.add_argument("--blur", type=float, default=0.0,
                        help="Optional Gaussian blur radius for mask smoothing in image mode")
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    args = parse_args()
    theme = THEMES.get(args.theme, THEMES["paper"]) 

    logging.info("Input: %s", args.input)
    logging.info("Output: %s", args.output)
    logging.info("Mode: %s", args.mode)
    logging.info("Theme: %s", theme.name)

    try:
        if args.mode == "vector":
            vector_mode_transform(args.input, args.output, theme)
        else:
            image_mode_transform(args.input, args.output, theme, dpi=args.dpi,
                                 threshold=args.threshold, blur_radius=args.blur)
    except Exception as e:
        logging.error("Conversion failed: %s", e)
        sys.exit(2)

    logging.info("Conversion completed successfully.")


if __name__ == "__main__":
    main()
