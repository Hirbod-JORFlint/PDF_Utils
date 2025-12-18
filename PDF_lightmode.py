#!/usr/bin/env python3
"""
pdf_lightmode_v1_fixed.py — v1.0

Improved PDF Light Mode Converter

- Primary (vector) mode: robust textual modification of content streams for
  common color operators (g/G, rg/RG, k/K, sc/SC/scn/SCN). This uses regex-based
  search/replace on decompressed content streams and prepends a solid background
  rectangle. This approach works reliably for PDFs where colors are set with
  literal numeric operands in content streams (very common).
- Fallback (image) mode: PyMuPDF + Pillow raster path for scanned/image PDFs.

Features implemented per spec:
- 3 themes (classic, warm, cool) as dataclasses with normalized RGB (0.0-1.0)
- Color detection (grayscale, RGB luminance, CMYK heuristic)
- CLI: input, output, --theme, --mode, --dpi, --threshold, --blur
- Progress bars via tqdm, good error handling, and informative console output.

Limitations:
- PDFs that set colors via indirect objects, named color spaces, or complex
  operators (color space names, indexed colors, patterns, shadings) may not
  be fully recolored by the regex pass. In those rare cases the image fallback
  will produce a correct visible result.

Run:
  python pdf_lightmode_v1_fixed.py in.pdf out.pdf --theme warm --mode vector
"""
from dataclasses import dataclass
import argparse
import sys
import io
import traceback
import math
import re
from typing import Tuple, List

try:
    import pikepdf
except Exception:
    pikepdf = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from PIL import Image, ImageFilter
except Exception:
    Image = None

import numpy as np
from tqdm import tqdm

__version__ = "v1.0"

# ---------------------------
# Theme dataclass & constants
# ---------------------------
@dataclass
class Theme:
    name: str
    bg_color: Tuple[float, float, float]  # normalized 0..1
    fg_color: Tuple[float, float, float]  # normalized 0..1


THEMES = {
    "classic": Theme("classic", (1.0, 1.0, 1.0), (0.0, 0.0, 0.0)),
    "warm": Theme("warm", (0.99, 0.97, 0.92), (0.20, 0.15, 0.10)),
    "cool": Theme("cool", (0.94, 0.97, 1.0), (0.10, 0.20, 0.35)),
}

# ---------------------------
# Utility functions
# ---------------------------
def clamp01(x: float) -> float:
    return max(0.0, min(1.0, float(x)))


def luminance_rgb_norm(r: float, g: float, b: float) -> float:
    return 0.299 * r + 0.587 * g + 0.114 * b


def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> Tuple[float, float, float]:
    # approximate conversion
    r = 1.0 - min(1.0, c + k)
    g = 1.0 - min(1.0, m + k)
    b = 1.0 - min(1.0, y + k)
    return (r, g, b)


def is_light_color_colorop(op: str, operands: List[float]) -> bool:
    """
    Decide whether a color is 'light' according to the rules from the spec.
    operands are normalized floats (0..1) or values that will be normalized inside.
    """
    # Normalize: if any value >1.5 assume 0..255 scale -> divide by 255
    nums = [float(x) for x in operands] if operands else []
    if any(abs(x) > 1.5 for x in nums):
        nums = [clamp01(x / 255.0) for x in nums]
    else:
        nums = [clamp01(x) for x in nums]

    op_l = op.lower()
    if op_l in ("g",):  # grayscale
        gray = nums[0] if nums else 1.0
        return gray >= 0.5
    elif op_l in ("rg",):
        if len(nums) < 3:
            # defensive fallback
            return False
        r, g, b = nums[:3]
        return luminance_rgb_norm(r, g, b) >= 0.5
    elif op_l in ("k",):
        # CMYK: K <= 0.5 and CMY sum <= 1.5 is light
        if len(nums) >= 4:
            c, m, y, k = nums[:4]
            return (k <= 0.5) and ((c + m + y) <= 1.5)
        else:
            # fallback: convert what we have to rgb and check luminance
            r, g, b = (nums + [0.0, 0.0, 0.0])[:3]
            return luminance_rgb_norm(r, g, b) >= 0.5
    elif op_l in ("sc", "scn", "scn"):  # generic color operators; try treated as RGB if 3 values
        if len(nums) >= 3:
            return luminance_rgb_norm(nums[0], nums[1], nums[2]) >= 0.5
        elif len(nums) == 1:
            return nums[0] >= 0.5
        else:
            return False
    else:
        return False


def rgb_to_operand_strings(rgb: Tuple[float, float, float]) -> List[str]:
    return [f"{clamp01(c):.4f}" for c in rgb]


def gray_from_rgb(rgb: Tuple[float, float, float]) -> str:
    g = luminance_rgb_norm(*rgb)
    return f"{clamp01(g):.4f}"


def cmyk_from_rgb(rgb: Tuple[float, float, float]) -> List[str]:
    r, g, b = rgb
    K = 1 - max(r, g, b)
    if K >= 1.0 - 1e-12:
        C = M = Y = 0.0
    else:
        denom = 1 - K
        C = (1 - r - K) / denom if denom != 0 else 0.0
        M = (1 - g - K) / denom if denom != 0 else 0.0
        Y = (1 - b - K) / denom if denom != 0 else 0.0
    return [f"{clamp01(C):.4f}", f"{clamp01(M):.4f}", f"{clamp01(Y):.4f}", f"{clamp01(K):.4f}"]


# ---------------------------
# Vector-mode: regex-based replacement
# ---------------------------
# numeric regex for PDF operand (int/float with optional exponent)
NUM_RE = r"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?"
# operators we will handle
OPS = r"(?:g|G|rg|RG|k|K|sc|SC|scn|SCN|SC|sC|sCn)"  # SC/special variants included conservatively

# Compile pattern: capture operand block followed by operator
# We will look for sequences where operands (1..4) precede the operator.
PATTERN = re.compile(
    rf"((?:{NUM_RE}\s+){{1,4}})({OPS})\b", flags=re.MULTILINE
)


def replace_colors_in_stream_textual(content_bytes: bytes, theme: Theme) -> bytes:
    """
    Replace color-setting numeric operators in the content stream bytes (decoded as latin1)
    with theme colors using the 'inverted logic' rule (source light -> theme.bg, else theme.fg).
    Returns new bytes.
    """
    try:
        s = content_bytes.decode("latin1")  # preserve bytes faithfully
    except Exception:
        s = content_bytes.decode("utf-8", errors="ignore")

    out_parts = []
    last_end = 0

    for m in PATTERN.finditer(s):
        start, end = m.span()
        operand_block, op = m.group(1), m.group(2)
        # copy preceding raw content
        out_parts.append(s[last_end:start])

        # parse operands
        operands = operand_block.strip().split()
        # convert to floats where possible
        nums = []
        for tok in operands:
            try:
                nums.append(float(tok))
            except Exception:
                nums.append(0.0)

        # determine if source color is light
        light = is_light_color_colorop(op, nums)

        target_rgb = theme.bg_color if light else theme.fg_color

        op_l = op.lower()
        if op_l == "g":
            # grayscale single operand
            new_operand = gray_from_rgb(target_rgb)
            replacement = f"{new_operand} {op}"
        elif op_l == "rg":
            rstr = " ".join(rgb_to_operand_strings(target_rgb))
            replacement = f"{rstr} {op}"
        elif op_l == "k":
            cmyk = cmyk_from_rgb(target_rgb)
            replacement = f"{' '.join(cmyk)} {op}"
        elif op_l in ("sc", "scn", "scn", "sc", "sc"):  # best-effort: if 3 operands -> rgb
            if len(nums) >= 3:
                rstr = " ".join(rgb_to_operand_strings(target_rgb))
                replacement = f"{rstr} {op}"
            elif len(nums) == 1:
                replacement = f"{gray_from_rgb(target_rgb)} {op}"
            else:
                # fallback to rgb
                replacement = f"{' '.join(rgb_to_operand_strings(target_rgb))} {op}"
        elif op_l in ("rg".lower(), "rg".upper()):
            rstr = " ".join(rgb_to_operand_strings(target_rgb))
            replacement = f"{rstr} {op}"
        elif op_l == "rg":  # redundant keep for clarity
            rstr = " ".join(rgb_to_operand_strings(target_rgb))
            replacement = f"{rstr} {op}"
        elif op_l == "k":
            cmyk = cmyk_from_rgb(target_rgb)
            replacement = f"{' '.join(cmyk)} {op}"
        elif op_l == "g":
            replacement = f"{gray_from_rgb(target_rgb)} {op}"
        else:
            # fallback
            replacement = f"{' '.join(rgb_to_operand_strings(target_rgb))} {op}"

        out_parts.append(replacement)
        last_end = end

    out_parts.append(s[last_end:])
    new_s = "".join(out_parts)
    return new_s.encode("latin1")


def create_solid_background_content(mediabox: Tuple[float, float, float, float], color: Tuple[float, float, float]) -> bytes:
    """
    Create a content stream that draws a solid rectangle covering the page with color (r,g,b normalized).
    Returns bytes (latin1).
    """
    r, g, b = color
    w = float(mediabox[2] - mediabox[0])
    h = float(mediabox[3] - mediabox[1])
    fmt = lambda x: f"{clamp01(x):.4f}"
    pieces = [
        "q",
        f"{fmt(r)} {fmt(g)} {fmt(b)} rg",
        f"0 0 {w:.4f} {h:.4f} re",
        "f",
        "Q",
    ]
    return ("\n".join(pieces) + "\n").encode("latin1")


def run_vector_mode(
    input_path: str,
    output_path: str,
    theme: Theme,
):
    """
    Vector-mode conversion using robust textual replacement of numeric color-setting
    operators, plus background rectangle insertion.
    """
    if pikepdf is None:
        raise RuntimeError("pikepdf is required for vector mode but is not installed.")

    print(f"[{__version__}] Running vector mode on '{input_path}' -> '{output_path}'")
    pdf = pikepdf.Pdf.open(input_path)
    pages = pdf.pages
    n = len(pages)
    print(f"PDF loaded: {n} pages. Theme={theme.name}")

    for i in tqdm(range(n), desc="Vector pages", unit="page"):
        page = pages[i]
        # determine mediabox (llx, lly, urx, ury)
        try:
            mx = page.MediaBox
            mediabox = (float(mx[0]), float(mx[1]), float(mx[2]), float(mx[3]))
        except Exception:
            try:
                cx = page.CropBox
                mediabox = (float(cx[0]), float(cx[1]), float(cx[2]), float(cx[3]))
            except Exception:
                mediabox = (0.0, 0.0, 612.0, 792.0)

        # gather raw content bytes (concatenate content streams if array)
        try:
            contents_obj = page.contents
            # page.contents may be a single Stream or an Array; pikepdf exposes iterable
            if isinstance(contents_obj, pikepdf.Array):
                raw_bytes = b"".join([c.read_bytes() for c in contents_obj])
            else:
                # single stream-like
                raw_bytes = contents_obj.read_bytes()
        except Exception:
            # safer fallback
            try:
                raw_bytes = page.obj.get("/Contents").read_bytes()
            except Exception:
                raw_bytes = b""

        if not raw_bytes:
            # nothing to change; still add background
            final_bytes = create_solid_background_content(mediabox, theme.bg_color)
            page.contents = pikepdf.Stream(pdf, final_bytes)
            continue

        # perform replacements
        try:
            new_bytes = replace_colors_in_stream_textual(raw_bytes, theme)
            # create final bytes as background cs + new content
            bg_bytes = create_solid_background_content(mediabox, theme.bg_color)
            final_bytes = bg_bytes + new_bytes
            page.contents = pikepdf.Stream(pdf, final_bytes)
        except Exception as e:
            # If replacement fails, still prepend background safely
            try:
                bg_bytes = create_solid_background_content(mediabox, theme.bg_color)
                # Prepend by creating a new array of streams if /Contents is array-compatible
                try:
                    existing = page.obj.get("/Contents")
                    bg_stream = pikepdf.Stream(pdf, bg_bytes)
                    if isinstance(existing, pikepdf.Array):
                        page.obj["/Contents"] = pikepdf.Array([bg_stream] + list(existing))
                    else:
                        page.obj["/Contents"] = pikepdf.Array([bg_stream, existing])
                except Exception:
                    page.contents = pikepdf.Stream(pdf, bg_bytes + raw_bytes)
            except Exception:
                print(f"Warning: could not update page {i}: {e}")

    # save output
    try:
        pdf.save(output_path)
        print(f"Saved vector-mode output to: {output_path}")
    except Exception as e:
        raise RuntimeError(f"Saving PDF failed: {e}")


# ---------------------------
# Image-mode: PyMuPDF + Pillow
# ---------------------------
def run_image_mode(
    input_path: str,
    output_path: str,
    theme: Theme,
    dpi: int = 150,
    threshold: int = 128,
    blur: float = 0.5,
):
    if fitz is None or Image is None:
        raise RuntimeError("PyMuPDF (fitz) and Pillow are required for image mode but are not available.")

    print(f"[{__version__}] Running image mode on '{input_path}' -> '{output_path}'")
    doc = fitz.open(input_path)
    new_doc = fitz.open()
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    for pnum in tqdm(range(len(doc)), desc="Raster pages", unit="page"):
        page = doc.load_page(pnum)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        arr = np.asarray(img).astype(np.float32) / 255.0
        lum = 0.299 * arr[..., 0] + 0.587 * arr[..., 1] + 0.114 * arr[..., 2]
        mask = lum < (threshold / 255.0)

        # blur mask for softness
        mask_img = Image.fromarray((mask * 255).astype(np.uint8)).convert("L")
        if blur and blur > 0:
            eff = blur * (dpi / 150.0)  # scale blur with dpi to keep consistent perception
            mask_img = mask_img.filter(ImageFilter.GaussianBlur(radius=eff))
        mask_arr = (np.asarray(mask_img).astype(np.float32) / 255.0)[..., None]

        bg_rgb = np.array(theme.bg_color).reshape((1, 1, 3))
        fg_rgb = np.array(theme.fg_color).reshape((1, 1, 3))

        # brighten non-text/image areas slightly
        bright = np.clip(arr * 1.1, 0.0, 1.0)

        # Compose: where mask (text) -> fg, else -> mix of bg and bright
        image_like = 1.0 - mask_arr
        non_text = (1.0 - image_like) * bg_rgb + image_like * bright
        result = mask_arr * fg_rgb + (1.0 - mask_arr) * non_text
        result = np.clip(result, 0.0, 1.0)

        out_img = Image.fromarray((result * 255).astype(np.uint8))

        # insert into new pdf page with same media box size in points
        rect = page.rect
        new_page = new_doc.new_page(width=rect.width, height=rect.height)
        # scale image bytes to page size
        img_bytes = io.BytesIO()
        out_img.save(img_bytes, format="PNG")
        img_bytes.seek(0)
        new_page.insert_image(fitz.Rect(0, 0, rect.width, rect.height), stream=img_bytes.getvalue(), keep_proportion=False)

    new_doc.save(output_path)
    print(f"Saved image-mode output to: {output_path}")


# ---------------------------
# CLI
# ---------------------------
def parse_args():
    p = argparse.ArgumentParser(description="PDF Light Mode Converter (v1.0 - fixed)")
    p.add_argument("input", help="Input PDF path")
    p.add_argument("output", help="Output PDF path")
    p.add_argument("--theme", choices=list(THEMES.keys()), default="classic", help="Theme: classic/warm/cool")
    p.add_argument("--mode", choices=["vector", "image"], default="vector", help="Mode: vector or image")
    p.add_argument("--dpi", type=int, default=150, help="DPI for image mode (default 150)")
    p.add_argument("--threshold", type=int, default=128, help="Text detection threshold (0-255) for image mode")
    p.add_argument("--blur", type=float, default=0.5, help="Mask blur radius for image mode")
    return p.parse_args()


def main():
    args = parse_args()
    theme = THEMES.get(args.theme, THEMES["classic"])

    print("=" * 70)
    print(f"PDF Light Mode Converter {__version__}")
    print(f"Input:  {args.input}")
    print(f"Output: {args.output}")
    print(f"Theme:  {theme.name} bg={theme.bg_color} fg={theme.fg_color}")
    print(f"Mode:   {args.mode}")
    if args.mode == "image":
        print(f"DPI: {args.dpi}  Threshold: {args.threshold}  Blur: {args.blur}")
    print("=" * 70)

    try:
        if args.mode == "vector":
            run_vector_mode(args.input, args.output, theme)
        else:
            run_image_mode(args.input, args.output, theme, dpi=args.dpi, threshold=args.threshold, blur=args.blur)
    except Exception as e:
        print("Error during processing:")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
