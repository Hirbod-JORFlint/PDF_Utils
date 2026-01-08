#!/usr/bin/env python3
"""
pdf_to_text_docx.py

Usage examples:
  python pdf_to_text_docx.py input.pdf --out docx --output out.docx
  python pdf_to_text_docx.py input.pdf --out txt --output out.txt --ocr
  python pdf_to_text_docx.py input.pdf --out docx --output out.docx --latex --ocr

Features:
 - Non-OCR path: extracts spans and attempts to preserve font name, size, bold/italic.
 - OCR path: rasterizes pages and runs Tesseract; more reliable for scanned PDFs but loses font metadata.
 - Exports optional LaTeX (.tex) best-effort.
 - Extracts images and inserts them into DOCX.
"""
import argparse
import os
import sys
import tempfile
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_BREAK
from PIL import Image
import io
import re

# optional import for OCR
try:
    import pytesseract
    HAVE_PYTESSERACT = True
except Exception:
    HAVE_PYTESSERACT = False

# ---------- Helpers ----------

def pil_from_pixmap(pix):
    """
    Convert a fitz.Pixmap to a PIL.Image.
    """
    mode = None
    if pix.n == 1:  # grayscale
        mode = "L"
        data = pix.samples
    elif pix.n == 3:
        mode = "RGB"
        data = pix.samples
    elif pix.n == 4:
        # contains alpha
        mode = "RGBA"
        data = pix.samples
    else:
        # fallback: convert to RGB via torgb
        pix = fitz.Pixmap(pix, 0)  # convert to RGB
        mode = "RGB"
        data = pix.samples
    img = Image.frombytes(mode, [pix.width, pix.height], data)
    return img

def clean_text_for_latex(s: str) -> str:
    """
    Escape LaTeX special characters (simple).
    """
    # order matters (escape backslash first)
    replacements = [
        ("\\", r"\textbackslash{}"),
        ("&", r"\&"),
        ("%", r"\%"),
        ("$", r"\$"),
        ("#", r"\#"),
        ("_", r"\_"),
        ("{", r"\{"),
        ("}", r"\}"),
        ("~", r"\textasciitilde{}"),
        ("^", r"\^{}"),
        ("<", r"\textless{}"),
        (">", r"\textgreater{}"),
    ]
    for a, b in replacements:
        s = s.replace(a, b)
    # preserve newlines
    return s

def infer_section_level_by_size(size_pt):
    """
    Heuristic to map font size to LaTeX sectioning level.
    (Very approximate; sizes vary.)
    """
    if size_pt is None:
        return None
    try:
        s = float(size_pt)
    except Exception:
        return None
    if s >= 20:
        return "section"
    if s >= 14:
        return "subsection"
    if s >= 11:
        return "subsubsection"
    return None

# ---------- Extraction ----------

def extract_structured_text_and_images(pdf_path, use_ocr=False, ocr_lang='eng', ocr_dpi=200, verbose=True):
    """
    Returns:
        pages: list of pages, each is a dict with:
            - 'spans': list of spans dicts: {'text','font','size','flags'}
            - 'images': list of image dicts: {'img_bytes', 'ext', 'xref', 'bbox'}
            - 'ocr_text' optional if use_ocr True
    """
    doc = fitz.open(pdf_path)
    pages_out = []

    for pno in range(len(doc)):
        page = doc[pno]
        page_dict = {"spans": [], "images": []}

        if use_ocr:
            # render page to image (pixmap) and run pytesseract
            if verbose:
                print(f"[OCR] Rendering page {pno+1}/{len(doc)} ...")
            mat = fitz.Matrix(ocr_dpi / 72, ocr_dpi / 72)  # scale
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = pil_from_pixmap(pix)
            if not HAVE_PYTESSERACT:
                raise RuntimeError("pytesseract not installed or not importable. Install pytesseract and the Tesseract binary.")
            txt = pytesseract.image_to_string(img, lang=ocr_lang)
            page_dict['ocr_text'] = txt
        else:
            # Use structural text extraction
            if verbose:
                print(f"[Extract] Page {pno+1}/{len(doc)} ...")
            blocks = page.get_text("dict")  # gives blocks -> lines -> spans
            # traverse blocks
            for b in blocks.get("blocks", []):
                if b.get("type", 0) == 0:  # text block
                    for line in b.get("lines", []):
                        for span in line.get("spans", []):
                            # span contains: text, font, size, flags, color, origin, bbox
                            s = {
                                "text": span.get("text", ""),
                                "font": span.get("font", ""),
                                "size": span.get("size", None),
                                "flags": span.get("flags", 0),
                                "bbox": span.get("bbox", None)
                            }
                            page_dict['spans'].append(s)
        # Extract images from page (always do this)
        image_list = page.get_images(full=True)
        if image_list:
            if verbose:
                print(f"  Found {len(image_list)} images on page {pno+1}")
        for imginfo in image_list:
            xref = imginfo[0]
            base_image = doc.extract_image(xref)
            img_bytes = base_image["image"]
            ext = base_image.get("ext", "png")
            bbox = None
            # bbox not directly from extract_image; could use imginfo tuple later for placement
            page_dict['images'].append({'img_bytes': img_bytes, 'ext': ext, 'xref': xref, 'bbox': bbox})
        pages_out.append(page_dict)
    doc.close()
    return pages_out

# ---------- Output writers ----------

def write_txt_from_pages(pages, output_path, use_ocr=False):
    with open(output_path, "w", encoding="utf-8") as f:
        for pno, page in enumerate(pages, start=1):
            f.write(f"\n\n--- Page {pno} ---\n\n")
            if use_ocr and 'ocr_text' in page:
                f.write(page['ocr_text'])
            else:
                # join spans in order (they are in reading order from PyMuPDF)
                for span in page['spans']:
                    f.write(span['text'])
            f.write("\n")
    print(f"Wrote TXT to {output_path}")

def write_docx_from_pages(pages, output_path, use_ocr=False, insert_page_breaks=True):
    doc = Document()
    # set default style's font if desired (leave default)
    for pno, page in enumerate(pages, start=1):
        if use_ocr and 'ocr_text' in page:
            # simply add OCR text; fix paragraphs by double newline
            text = page['ocr_text']
            for para in text.splitlines():
                doc.add_paragraph(para)
        else:
            # structured spans. We'll group consecutive spans separated by newline-like spans into a paragraph.
            # Simplest: create a paragraph per span but try to preserve runs
            p = None
            for span in page['spans']:
                txt = span['text']
                if not txt:
                    continue
                # new paragraph when text contains newline
                if p is None:
                    p = doc.add_paragraph()
                run = p.add_run(txt)
                # font settings best-effort
                try:
                    rfont = run.font
                except Exception:
                    rfont = run.font
                # font name
                fontname = span.get('font')
                if fontname:
                    # PyMuPDF font names can include encoding / subsetting; attempt to clean
                    clean_font = re.sub(r"[-,]._.*", "", fontname)
                    rfont.name = clean_font
                # size
                size = span.get('size')
                if size:
                    try:
                        rfont.size = Pt(float(size))
                    except Exception:
                        pass
                # bold/italic heuristics
                flags = span.get('flags', 0)
                # Heuristic: font name contains 'Bold' or 'Italic' or flags bitmask
                if fontname and ("Bold" in fontname or "bold" in fontname):
                    run.bold = True
                if fontname and ("Italic" in fontname or "Oblique" in fontname or "italic" in fontname):
                    run.italic = True
                # flags: according to MuPDF, flags may include 2 for bold (not guaranteed)
                try:
                    if flags & 2:
                        run.bold = True
                    if flags & 1:
                        run.italic = True
                except Exception:
                    pass
            # end page spans
        # insert images for page (append at end of page)
        for imgdict in page.get('images', []):
            img_bytes = imgdict['img_bytes']
            ext = imgdict.get('ext', 'png').lower()
            # write to temp file and add
            with tempfile.NamedTemporaryFile(delete=False, suffix='.' + ext) as tf:
                tf.write(img_bytes)
                tmpfn = tf.name
            try:
                # limit size to page width reasonably
                doc.add_picture(tmpfn, width=Inches(6))  # best-effort
            except Exception:
                # fallback: try no size
                try:
                    doc.add_picture(tmpfn)
                except Exception:
                    pass
            finally:
                try:
                    os.unlink(tmpfn)
                except Exception:
                    pass
        if insert_page_breaks and pno < len(pages):
            doc.add_page_break()
    doc.save(output_path)
    print(f"Wrote DOCX to {output_path}")

def write_latex_from_pages(pages, output_path, use_ocr=False):
    lines = []
    lines.append(r"\documentclass{article}")
    lines.append(r"\usepackage[utf8]{inputenc}")
    lines.append(r"\usepackage{graphicx}")
    lines.append(r"\begin{document}")
    for pno, page in enumerate(pages, start=1):
        lines.append(f"% --- Page {pno} ---")
        if use_ocr and 'ocr_text' in page:
            txt = page['ocr_text']
            for para in txt.splitlines():
                if para.strip() == "":
                    lines.append("")
                else:
                    lines.append(clean_text_for_latex(para) + r"\\")
        else:
            # try to detect headings by font size
            last_section = None
            for span in page['spans']:
                text = span.get('text', '')
                if not text.strip():
                    continue
                size = span.get('size')
                sect = infer_section_level_by_size(size)
                if sect and len(text.strip()) > 2:
                    if sect != last_section:
                        lines.append(r"\{}{{{}}}".format(sect, clean_text_for_latex(text.strip())))
                        last_section = sect
                        continue
                # otherwise add as paragraph
                lines.append(clean_text_for_latex(text.strip()) + r"\\")
        # images: reference as comments (embedding actual images in LaTeX requires saving files)
        if page.get('images'):
            for idx, imgdict in enumerate(page.get('images')):
                # save image next to tex file (name it)
                ext = imgdict.get('ext', 'png').lower()
                img_fname = f"page{pno}_img{idx + 1}.{ext}"
                lines.append(r"\begin{figure}[h]")
                lines.append(r"\centering")
                lines.append(r"\includegraphics[width=0.8\linewidth]{" + img_fname + r"}")
                lines.append(r"\end{figure}")
        lines.append(r"\newpage")
    lines.append(r"\end{document}")
    # write
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Wrote LaTeX to {output_path}")
    # Also write image files next to it for user's convenience
    base_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    for pno, page in enumerate(pages, start=1):
        for idx, imgdict in enumerate(page.get('images', [])):
            ext = imgdict.get('ext', 'png').lower()
            img_fname = os.path.join(base_dir, f"page{pno}_img{idx + 1}.{ext}")
            with open(img_fname, "wb") as g:
                g.write(imgdict['img_bytes'])
    if any(p.get('images') for p in pages):
        print("Saved extracted images next to the .tex file (filenames like page<N>_img<M>.<ext>)")

# ---------- CLI ----------

def parse_args():
    p = argparse.ArgumentParser(description="Convert PDF -> DOCX/TXT (with optional OCR and LaTeX export).")
    p.add_argument("input", help="Input PDF file path.")
    p.add_argument("--out", choices=["docx", "txt", "both"], default="docx", help="Which output text format to write.")
    p.add_argument("--output", "-o", help="Output file path. If omitted, derived from input. For --out both, output is used as stem.")
    p.add_argument("--ocr", action="store_true", help="Force OCR (using Tesseract). Useful for scanned PDFs.")
    p.add_argument("--latex", action="store_true", help="Also output a LaTeX (.tex) file.")
    p.add_argument("--lang", default="eng", help="Language for Tesseract OCR (default: eng).")
    p.add_argument("--dpi", type=int, default=200, help="DPI for rasterization for OCR (default 200). Higher = better OCR but slower.")
    p.add_argument("--no-page-breaks", dest="page_breaks", action="store_false", help="Don't insert page breaks in DOCX.")
    p.add_argument("--verbose", action="store_true", help="Verbose output.")
    return p.parse_args()

def main():
    args = parse_args()
    inp = args.input
    if not os.path.exists(inp):
        print("Input file not found:", inp, file=sys.stderr)
        sys.exit(2)
    # determine output paths
    stem = os.path.splitext(os.path.basename(inp))[0]
    out_spec = args.output
    if not out_spec:
        if args.out == "docx":
            out_spec = stem + ".docx"
        elif args.out == "txt":
            out_spec = stem + ".txt"
        else:
            out_spec = stem  # used as stem for both
    # Extract
    pages = extract_structured_text_and_images(inp, use_ocr=args.ocr, ocr_lang=args.lang, ocr_dpi=args.dpi, verbose=args.verbose)

    if args.out in ("txt", "both"):
        if args.out == "both":
            out_txt = out_spec if out_spec.endswith(".txt") else out_spec + ".txt"
        else:
            out_txt = out_spec
        write_txt_from_pages(pages, out_txt, use_ocr=args.ocr)

    if args.out in ("docx", "both"):
        if args.out == "both":
            out_docx = out_spec if out_spec.endswith(".docx") else out_spec + ".docx"
        else:
            out_docx = out_spec
        write_docx_from_pages(pages, out_docx, use_ocr=args.ocr, insert_page_breaks=args.page_breaks)

    if args.latex:
        out_tex = os.path.splitext(out_spec)[0] + ".tex"
        write_latex_from_pages(pages, out_tex, use_ocr=args.ocr)

if __name__ == "__main__":
    main()
