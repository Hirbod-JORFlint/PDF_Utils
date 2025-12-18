#!/usr/bin/env python3
"""
pdf_compressor.py

Usage:
    python pdf_compressor.py input.pdf output.pdf level(1-4) [--force-scanned] [--keep-temporary]

Levels:
    1 - Low (lossless / gentle)
    2 - Medium (balanced)
    3 - High (aggressive)
    4 - Extreme (very aggressive, rasterization + heavy downsampling for scanned docs)

Notes:
    - Requires pikepdf, Pillow, PyPDF2.
    - Ghostscript is optional but recommended (used for extreme compression and scanned optimizations).
"""
import os
import sys
import shutil
import tempfile
import subprocess
from io import BytesIO

from PIL import Image
import pikepdf
from PyPDF2 import PdfReader

# ---------- SETTINGS / TUNABLES ----------
# Image quality / size settings per level (these will be used when recompressing images via Pillow)
IMAGE_SETTINGS = {
    1: dict(max_size=(4000, 4000), quality=92, progressive=True, grayscale=False),
    2: dict(max_size=(2500, 2500), quality=78, progressive=True, grayscale=False),
    3: dict(max_size=(1600, 1600), quality=60, progressive=True, grayscale=False),
    4: dict(max_size=(1000, 1000), quality=30, progressive=True, grayscale=True),  # extreme: smaller + grayscale
}

# Threshold to decide "scanned" based on average extracted characters per page
SCANNED_TEXT_THRESHOLD = 40  # if average chars/page < this, treat as scanned

# Ghostscript device mappings for levels (fallback)
GS_PDFSETTINGS = {
    1: "/prepress",
    2: "/printer",
    3: "/ebook",
    4: "/screen",
}


# ---------- Helper functions ----------
def safe_pdf_save(pdf, output_path):
    """
    Save PDF with best available compression options depending on pikepdf version.
    """
    save_kwargs = {
        "compress_streams": True,
        "object_stream_mode": pikepdf.ObjectStreamMode.generate,
    }

    # optimize_streams exists only in newer pikepdf
    if "optimize_streams" in pdf.save.__code__.co_varnames:
        save_kwargs["optimize_streams"] = True

    pdf.save(output_path, **save_kwargs)

def extract_text_density(input_pdf, max_pages=20):
    """
    Measure average characters per page by sampling up to max_pages pages.
    Returns (avg_chars_per_page, total_pages_sampled).
    """
    try:
        reader = PdfReader(input_pdf)
        n_pages = len(reader.pages)
        sample_count = min(n_pages, max_pages)
        total_chars = 0
        pages_to_sample = range(sample_count)
        for i in pages_to_sample:
            page = reader.pages[i]
            text = page.extract_text()
            total_chars += len(text) if text else 0
        avg = total_chars / sample_count if sample_count else 0
        return avg, n_pages
    except Exception:
        # If text extraction fails, assume scanned
        return 0, 0


def is_scanned_pdf(input_pdf):
    avg_chars, total_pages = extract_text_density(input_pdf)
    # If average characters per page are lower than threshold, treat as scanned
    scanned = avg_chars < SCANNED_TEXT_THRESHOLD
    # Debug info
    print(f"[detect] average chars/page={avg_chars:.1f} over {min(total_pages,20)} sampled -> scanned={scanned}")
    return scanned


def compress_image_bytes(img_bytes, level, for_scanned=False):
    """
    Recompress image bytes using Pillow according to the level settings.
    Returns new image bytes suitable for inserting into PDF (JPEG encoded).
    """
    settings = IMAGE_SETTINGS[level]
    try:
        img = Image.open(BytesIO(img_bytes))
    except Exception:
        # If PIL cannot open it, return original
        return img_bytes

    # Convert indexed/alpha images to RGB (or L for grayscale)
    if img.mode in ("P", "RGBA", "LA"):
        if settings.get("grayscale") or for_scanned:
            img = img.convert("L")
        else:
            img = img.convert("RGB")
    elif img.mode == "CMYK":
        # convert to RGB first
        img = img.convert("RGB")
    elif settings.get("grayscale") or for_scanned:
        # If grayscale requested, convert
        img = img.convert("L")

    # Downsample by resizing to within max_size, preserving aspect ratio
    max_w, max_h = settings["max_size"]
    img.thumbnail((max_w, max_h), Image.LANCZOS)

    out = BytesIO()
    save_kwargs = {"format": "JPEG", "quality": settings["quality"], "optimize": True}
    # progressive is supported for RGB; for grayscale Pillow will handle textures.
    if settings.get("progressive"):
        save_kwargs["progressive"] = True

    try:
        img.save(out, **save_kwargs)
        return out.getvalue()
    except Exception:
        # If JPEG saving fails (rare), fallback to original bytes
        return img_bytes


def replace_images_in_pdf(input_pdf, output_pdf, level, treat_scanned=False):
    with pikepdf.open(input_pdf) as pdf:
        for page in pdf.pages:
            try:
                images = page.images
            except Exception:
                images = {}

            for name, image_obj in list(images.items()):
                try:
                    raw = image_obj.read_bytes()
                    new_bytes = compress_image_bytes(raw, level, for_scanned=treat_scanned)
                    if new_bytes and len(new_bytes) < len(raw) - 16:
                        stream = pikepdf.Stream(
                            pdf,
                            new_bytes,
                            Filter=pikepdf.Name("/DCTDecode")
                        )
                        page.images[name] = stream
                except Exception:
                    continue

        safe_pdf_save(pdf, output_pdf)



def run_ghostscript(input_pdf, output_pdf, level, scanned=False, extreme=False):
    """
    Call Ghostscript to further compress. We choose settings depending on level / scanned.
    For extreme/scanned we may apply explicit downsampling/resampling options.
    """
    gs_setting = GS_PDFSETTINGS.get(level, "/screen")
    # Base command
    cmd = [
        "gs",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={gs_setting}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
    ]

    # If scanned or extreme, add explicit downsampling and grayscale options
    if scanned or extreme or level == 4:
        # Aggressive downsample and force grayscale if extreme
        # Keep these flags conservative but impactful:
        cmd += [
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Subsample",
            f"-dColorImageResolution=100",
            f"-dGrayImageResolution=100",
            f"-dMonoImageResolution=100",
            "-dAutoFilterColorImages=false",
            "-dAutoFilterGrayImages=false",
            "-dAutoFilterMonoImages=false",
            "-dEncodeColorImages=true",
            "-dEncodeGrayImages=true",
            "-dEncodeMonoImages=true",
        ]
        if extreme or (scanned and level >= 3):
            # More aggressive
            cmd += [
                "-dColorImageResolution=72",
                "-dGrayImageResolution=72",
                "-dMonoImageResolution=72",
            ]
        if extreme:
            # Force grayscale output to shrink size further
            cmd += ["-sColorConversionStrategy=Gray", "-dConvertCMYKImagesToRGB=false"]

    cmd += [f"-sOutputFile={output_pdf}", input_pdf]

    # Execute
    try:
        subprocess.run(cmd, check=True)
        return True
    except FileNotFoundError:
        print("[gs] Ghostscript (gs) not found on PATH; skipping Ghostscript stage.")
        return False
    except subprocess.CalledProcessError as e:
        print(f"[gs] Ghostscript failed: {e}")
        return False


# ---------- Main high-level API ----------
def compress_pdf(input_pdf, output_pdf, level, force_scanned=False, keep_temp=False):
    assert level in (1, 2, 3, 4), "level must be 1..4"
    tmpdir = tempfile.mkdtemp(prefix="pdfcmp_")
    try:
        base_temp = os.path.join(tmpdir, "step0.pdf")
        shutil.copyfile(input_pdf, base_temp)

        scanned = force_scanned or is_scanned_pdf(input_pdf)

        # Stage 1: attempt to replace embedded images with recompressed versions
        stage1 = os.path.join(tmpdir, "stage1.pdf")
        print(f"[stage] recompressing embedded images (level {level}) ...")
        replaced = replace_images_in_pdf(base_temp, stage1, level, treat_scanned=scanned)

        # If nothing was replaced, still copy forward
        if not os.path.exists(stage1):
            shutil.copyfile(base_temp, stage1)

        # Stage 2: structural optimization with pikepdf (even if no images changed)
        stage2 = os.path.join(tmpdir, "stage2.pdf")
        print("[stage] optimizing PDF structure (pikepdf)...")
        with pikepdf.open(stage1) as pdf:
            safe_pdf_save(pdf, stage2)

        final_temp = os.path.join(tmpdir, "final.pdf")

        # Stage 3: conditional Ghostscript pass for stronger compression
        # Rules:
        #   - If level == 4 (extreme) -> run Ghostscript with aggressive settings
        #   - If scanned and level >=3 -> run Ghostscript too
        use_gs = False
        if level == 4:
            use_gs = True
            extreme = True
        elif scanned and level >= 3:
            use_gs = True
            extreme = False
        else:
            use_gs = False
            extreme = False

        if use_gs:
            print("[stage] running Ghostscript heavy pass...")
            gs_ok = run_ghostscript(stage2, final_temp, level, scanned=scanned, extreme=extreme)
            if not gs_ok:
                print("[stage] Ghostscript failed or missing: falling back to pikepdf-only result.")
                shutil.copyfile(stage2, final_temp)
        else:
            print("[stage] skipping Ghostscript pass (not needed for this level).")
            shutil.copyfile(stage2, final_temp)

        # Move final result
        shutil.copyfile(final_temp, output_pdf)
        print(f"[done] saved compressed PDF to: {output_pdf}")

    finally:
        if keep_temp:
            print(f"[temp] temporary files kept in: {tmpdir}")
        else:
            shutil.rmtree(tmpdir, ignore_errors=True)


# ---------- CLI ----------
def print_usage_and_exit():
    print("Usage: python pdf_compressor.py input.pdf output.pdf level(1-4) [--force-scanned] [--keep-temporary]")
    sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print_usage_and_exit()

    inp = sys.argv[1]
    out = sys.argv[2]
    try:
        lvl = int(sys.argv[3])
    except ValueError:
        print("Level must be integer 1..4")
        sys.exit(1)

    force_scanned = "--force-scanned" in sys.argv[4:]
    keep_temp = "--keep-temporary" in sys.argv[4:]

    compress_pdf(inp, out, lvl, force_scanned=force_scanned, keep_temp=keep_temp)
