#!/usr/bin/env python3
"""
pdf_to_text_docx_high_fidelity.py

High-Fidelity PDF -> DOCX Converter
Improvements over original:
 - Matches DOCX page size to PDF page size.
 - Intelligent Font Name mapping (PDF internal names -> Windows/Word names).
 - Hyperlink detection and embedding.
 - Superscript detection.
 - Zero-spacing layout (prevents large gaps between lines).
 - In-memory image processing (no temp files).
Usage examples:
  python pdf_to_text_docx_with_columns.py input.pdf --out docx --output out.docx
  python pdf_to_text_docx_with_columns.py input.pdf --out txt --output out.txt --ocr
  python pdf_to_text_docx_with_columns.py input.pdf --out docx --latex --max-columns 3 --col-gap 0.18
Usage:
  python pdf_to_text_docx_high_fidelity.py input.pdf --output out.docx
"""

import argparse
import os
import sys
import re
import io

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==========================================
# Helpers
# ==========================================

def int_to_rgb(color_value):
    """Convert PyMuPDF color int/tuple to RGB tuple."""
    if isinstance(color_value, int):
        # Some PDFs store as simple int
        r = (color_value >> 16) & 255
        g = (color_value >> 8) & 255
        b = color_value & 255
        return (r, g, b)
    elif isinstance(color_value, (list, tuple)) and len(color_value) >= 3:
        # PyMuPDF floats 0.0-1.0 or ints 0-255
        if all(isinstance(c, float) for c in color_value):
            return (int(color_value[0]*255), int(color_value[1]*255), int(color_value[2]*255))
        return tuple(map(int, color_value[:3]))
    return (0, 0, 0)

def map_font_name(fname, flags=None):
    """
    Map PDF font names to standard Word fonts to improve visual fidelity.
    - Normalizes subset prefixes (AAAAAA+FontName), suffix tokens, and casing.
    - Tries to preserve family (serif/sans/mono) and common names.
    """
    if not fname:
        return "Calibri"

    # Remove subset prefix like "ABCDEE+FontName" and any encoding suffix
    fname = fname.split('+')[-1]
    fname = fname.split('-')[0]

    # Clean common suffix tokens and punctuation
    fname_clean = re.sub(r'^[A-Z]{6}\+', '', fname)
    fname_clean = re.sub(r'[-_,.](bold|italic|regular|mt|ps|std|roman|oblique|semi|condensed|narrow)$',
                        '', fname_clean, flags=re.I).strip()
    fname_lower = fname_clean.lower()

    # family heuristics
    if "times" in fname_lower or "serif" in fname_lower:
        return "Times New Roman"
    if "arial" in fname_lower or "helvetica" in fname_lower:
        return "Arial"
    if "courier" in fname_lower or "mono" in fname_lower or "monospace" in fname_lower:
        return "Courier New"
    if "calibri" in fname_lower:
        return "Calibri"
    if "cambria" in fname_lower:
        return "Cambria"
    if "georgia" in fname_lower:
        return "Georgia"

    # Fallback:
    return fname_clean or "Calibri"

def add_hyperlink(paragraph, url, text, color, is_bold, is_italic, font_name, font_size):
    """
    Manually inject a hyperlink into the DOCX XML.
    python-docx does not natively support adding links to runs easily.
    """
    # This gets the paragraph's relation part ID
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)

    # Create the <w:hyperlink> tag
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create the <w:r> (run) tag
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Color (usually blue for links, but we use source color)
    c_el = OxmlElement('w:color')
    c_el.set(qn('w:val'), "{:02x}{:02x}{:02x}".format(*color))
    rPr.append(c_el)

    # Underline (standard for links)
    u_el = OxmlElement('w:u')
    u_el.set(qn('w:val'), 'single')
    rPr.append(u_el)

    # Fonts/Style
    if is_bold:
        rPr.append(OxmlElement('w:b'))
    if is_italic:
        rPr.append(OxmlElement('w:i'))
    
    # Text
    text_el = OxmlElement('w:t')
    text_el.text = text
    
    run.append(rPr)
    run.append(text_el)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

def is_box_inside(inner, outer):
    ix0, iy0, ix1, iy1 = inner
    ox0, oy0, ox1, oy1 = outer
    cx = (ix0 + ix1) / 2
    cy = (iy0 + iy1) / 2
    return (ox0 <= cx <= ox1) and (oy0 <= cy <= oy1)

# ==========================================
# Extraction Logic
# ==========================================

def get_link_target(bbox, links):
    """Find if a bbox intersects with a known link area."""
    # links is a list of dicts from page.get_links()
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    
    for link in links:
        lb = link['from'] # Rect
        if lb.x0 <= cx <= lb.x1 and lb.y0 <= cy <= lb.y1:
            return link.get('uri')
    return None

def analyze_columns_simple(spans, page_width):
    """
    Simple heuristic to detect if text belongs to left or right column
    to calculate indentation correctly.
    """
    mid = page_width / 2
    # Simple cluster check: how many spans are strictly left vs strictly right?
    # This is a basic check.
    return [(0, page_width)] 

def extract_page_content(page):
    elements = []
    links = page.get_links()

    # 1. Tables
    table_rects = []
    tables = page.find_tables()
    for tab in tables:
        bbox = tab.bbox
        table_rects.append(bbox)
        elements.append({
            'type': 'table',
            'bbox': bbox,
            'y_sort': bbox[1],
            'data': tab.extract()
        })

    # 2. Text & Images
    text_blocks = page.get_text("dict", sort=True)["blocks"]

    page_height = page.rect.height
    header_threshold = page_height * 0.08  # Top 8%
    footer_threshold = page_height * 0.92  # Bottom 8%
    
    for b in text_blocks:
        bbox = b['bbox']
        # Identify if block is a Header/Footer
        is_margin_content = bbox[3] < header_threshold or bbox[1] > footer_threshold
        
        # Optional: Skip common artifacts like single-digit page numbers in margins
        if is_margin_content and len(b.get("lines", [])) == 1:
            if re.match(r'^\d+$', page.get_text("text", clip=bbox).strip()):
                continue
            
        if any(is_box_inside(bbox, tr) for tr in table_rects):
            continue

        if b['type'] == 0:  # Text
            para_lines = []
            for line in b['lines']:
                line_spans = []
                for span in line['spans']:
                    uri = get_link_target(span['bbox'], links)
                    line_spans.append({
                        'text': span['text'],
                        'font': span['font'],
                        'size': span['size'],
                        'color': span['color'],
                        'flags': span['flags'],
                        'bbox': span['bbox'],
                        'origin': span['origin'],
                        'link': uri
                    })
                if line_spans:
                    para_lines.append({'bbox': line['bbox'], 'spans': line_spans})
            
            if para_lines:
                paragraphs = []
                def avg_font_size(line):
                    sizes = [s.get('size') for s in line['spans'] if s.get('size')]
                    return float(sizes[0]) if sizes else 12.0

                current = {'bbox': para_lines[0]['bbox'], 'spans': list(para_lines[0]['spans'])}
                prev_center_y = (para_lines[0]['bbox'][1] + para_lines[0]['bbox'][3]) / 2.0
                current_font = avg_font_size(para_lines[0])

                for ln in para_lines[1:]:
                    center_y = (ln['bbox'][1] + ln['bbox'][3]) / 2.0
                    gap = abs(center_y - prev_center_y)
                    # previous: threshold = max(6.0, 0.75 * current_font)
                    # Improved: use both an absolute minimum and a fraction of the font size,
                    # but use the actual bbox vertical gap to decide. This reduces accidental merges.
                    # current_font is in points.
                    min_gap_pts = 3.0  # don't treat sub-3pt differences as same-line
                    rel_gap_factor = 0.45  # fraction of font size to accept as same-line
                    threshold = max(min_gap_pts, rel_gap_factor * current_font)

                    # Check vertical gap AND horizontal alignment (left margin)
                    x_offset = abs(ln['bbox'][0] - current['bbox'][0])
                    if gap <= threshold and x_offset < 10:
                        last_span = current['spans'][-1]
                        if last_span['text'].rstrip().endswith('-') and len(last_span['text'].strip()) > 1:
                            last_span['text'] = last_span['text'].rstrip().rstrip('-')
                        else:
                            if not last_span['text'].endswith(' '):
                                last_span['text'] += ' '
                        current['spans'].extend(ln['spans'])
                        c = list(current['bbox'])
                        c[2] = max(c[2], ln['bbox'][2])
                        c[3] = max(c[3], ln['bbox'][3])
                        current['bbox'] = tuple(c)
                        prev_center_y = center_y
                    else:
                        paragraphs.append(current)
                        current = {'bbox': ln['bbox'], 'spans': list(ln['spans'])}
                        prev_center_y = center_y
                        current_font = avg_font_size(ln)

                paragraphs.append(current)

                for para in paragraphs:
                    elements.append({
                        'type': 'text',
                        'bbox': para['bbox'],
                        'y_sort': para['bbox'][1],
                        'lines': [para],
                        'indent_pt': para['bbox'][0]
                    })
                    
        elif b['type'] == 1:  # Image
            img_bytes = b.get("image")
            if img_bytes:
                elements.append({
                    'type': 'image',
                    'bbox': bbox,
                    'y_sort': bbox[1],
                    'bytes': img_bytes,
                    'ext': b.get("ext", "png")
                })

    # Detect logical columns based on the distribution of x0 coordinates
    x_coords = sorted([el['bbox'][0] for el in elements])
    columns = [0]
    # Detect logical columns based on the distribution of x0 coordinates
    # Filter out very small overlaps and group x-coordinates into clusters
    if x_coords:
        # Use a 10% page width threshold to identify distinct column starts
        col_threshold = page.rect.width * 0.10
        for i in range(1, len(x_coords)):
            if x_coords[i] - columns[-1] > col_threshold:
                columns.append(x_coords[i])                           
    # Sort elements primarily by column (x-coordinate), then by row (y-coordinate)
    # This ensures that in multi-column layouts, the left column is read fully before the right.
    def get_reading_order(el):
        bbox = el['bbox']
        # Snap x-coordinate to the nearest column start to handle slight misalignments
        col_idx = 0
        for idx, col_start in enumerate(columns):
            if bbox[0] >= col_start - 5: # 5pt tolerance
                col_idx = idx
        # Sort by Column first, then Y (top-to-bottom), then X (left-to-right)
        return (col_idx, round(bbox[1], 1), round(bbox[0], 1))
    
    # New Logic:
    # 1. 'elements' currently contains tables and images (processed first).
    # 2. Add processed text paragraphs to 'elements'.
    # 3. Sort primarily by vertical position, BUT with a column-aware grouping 
    #    or simply trust the "sort=True" block order if you remove the manual sort.
    
    # Robust Fix: Sort by vertical bands to keep columns distinct
    # (Groups items within 50px vertical distinctness, then sorts Left-to-Right)
    elements.sort(key=get_reading_order)
    
    return elements

# ==========================================
# DOCX Writing
# ==========================================

def write_to_docx(doc, elements, page_width, page_height):
    
    # Match Page Size logic
    section = doc.sections[-1]
    # page_width/page_height are PDF points (72 pt == 1 inch).
    # Convert to inches so python-docx receives the intended size.
    section.page_width = Inches(page_width / 72.0)
    section.page_height = Inches(page_height / 72.0)
    
    # Reduce margins to allow PDF-like absolute positioning approximations
    # We use a small left margin and control the rest via indentation
    margin_buffer = 36 # 0.5 inch
    section.left_margin = Pt(margin_buffer)
    section.right_margin = Pt(margin_buffer)
    section.top_margin = Pt(margin_buffer)
    section.bottom_margin = Pt(margin_buffer)

    for el in elements:
        
        if el['type'] == 'table':
            data = el['data']
            if not data or not data[0]: continue
            
            rows = len(data)
            cols = len(data[0])
            table = doc.add_table(rows=rows, cols=cols)
            table.style = 'Table Grid'
            table.autofit = False 
            # Calculate proportional widths from the PDF bbox
            total_pts = el['bbox'][2] - el['bbox'][0]
            col_width = Inches((total_pts / 72.0) / cols)
            for column in table.columns:
                column.width = col_width
            
            for r, row_data in enumerate(data):
                row_cells = table.rows[r].cells
                for c, cell_text in enumerate(row_data):
                    if c >= len(row_cells): break
                    cell = row_cells[c]
                    t_val = (cell_text or "").strip()
                    
                    # Create a run to allow formatting
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if r == 0 else WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(t_val)
                    if r == 0: # Bold headers
                        run.bold = True
                        cell.shading.background_pattern_color = "EDEDED"
            
            doc.add_paragraph() # spacer

        elif el['type'] == 'image':
            img_stream = io.BytesIO(el['bytes'])
            try:
                # Calculate width in inches based on PDF points (1/72 inch)
                # Cap it at page width - margins
                img_w_pts = el['bbox'][2] - el['bbox'][0]
                img_w_inches = img_w_pts / 72.0
                
                doc.add_picture(img_stream, width=Inches(img_w_inches))
                last_p = doc.paragraphs[-1]
                
                # Try to approximate alignment
                page_mid = page_width / 2
                img_mid = (el['bbox'][0] + el['bbox'][2]) / 2
                
                if abs(img_mid - page_mid) < 20:
                    last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif img_mid > page_mid:
                    last_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    last_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
            except Exception as e:
                print(f"Warning: Image skipped: {e}")

        elif el['type'] == 'text':
            p = doc.add_paragraph()
            
            # Get full text to check for list patterns
            full_text = "".join([s['text'] for l in el['lines'] for s in l['spans']]).strip()

            # Expanded regex for nested numbers (1.1) and varied bullet glyphs
            if re.match(r'^([\u2022\u2023\u25E6\u2043\u27A2\-\*])\s+', full_text):
                p.style = 'List Bullet'
                pf.left_indent = Pt(0) # Let Word style handle the bullet indent
            elif re.match(r'^(\d+(\.\d+)*|[A-Za-z]\.)\s+', full_text):
                p.style = 'List Number'
                pf.left_indent = Pt(0)
    
            pf = p.paragraph_format
            # Indentation
            raw_indent = el['indent_pt']
            relative_indent = max(0, raw_indent - margin_buffer)
            pf.left_indent = Pt(relative_indent)
            
            # Smart spacing: larger gap for headings, standard gap for body text
            if p.style.name.startswith('Heading') or p.style.name == 'Title':
                pf.space_before = Pt(12)
                pf.space_after = Pt(6)
            else:
                pf.space_after = Pt(8) # Standard paragraph spacing

            # Heuristic: Calculate max font size in this paragraph
            p_max_size = 0
            for line in el['lines']:
                for span in line['spans']:
                    s = span.get('size', 0)
                    if s > p_max_size: p_max_size = s
            
            # Enhanced: Use size + bold flags for semantic hierarchy
            is_mostly_bold = any(bool(s.get('flags', 0) & 16) for l in el['lines'] for s in l['spans'])
            
            # Calculate base size from all text elements to find "Normal"
            all_sizes = [s.get('size', 12) for el in elements if el['type'] == 'text' for ln in el['lines'] for s in ln['spans']]
            base_size = max(all_sizes, key=all_sizes.count) if all_sizes else 12

            # Apply semantic styles based on size relative to the most common font size
            if p_max_size >= base_size * 1.5:
                p.style = 'Title'
            elif p_max_size >= base_size + 4:
                p.style = 'Heading 1'
            elif p_max_size >= base_size + 2:
                p.style = 'Heading 2'
            elif is_mostly_bold and p_max_size >= base_size:
                p.style = 'Heading 3'
            
            for line in el['lines']:
                spans = line['spans']
                for si, span in enumerate(spans):
                    text = span.get('text') or ''
                    if not text:
                        continue

                    # If there is an explicit link in this span, keep previous hyperlink logic
                    if span.get('link'):
                        rgb = int_to_rgb(span['color'])
                        is_bold = bool(span['flags'] & 16)
                        is_italic = bool(span['flags'] & 2)
                        fname = map_font_name(span.get('font'), span.get('flags'))
                        add_hyperlink(p, span['link'], text, rgb, is_bold, is_italic, fname, span.get('size'))
                        # continue to possible spacing logic below (we still may want a space after the link)
                    else:
                        run = p.add_run(text)

                        # font name + rFonts xml
                        font_name = map_font_name(span.get('font'))
                        run.font.name = font_name
                        try:
                            rfonts = OxmlElement('w:rFonts')
                            rfonts.set(qn('w:ascii'), font_name)
                            rfonts.set(qn('w:hAnsi'), font_name)
                            rfonts.set(qn('w:eastAsia'), font_name)
                            rPr = run._element.get_or_add_rPr()
                            rPr.append(rfonts)
                        except Exception:
                            pass

                        # size
                        try:
                            if span.get('size'):
                                run.font.size = Pt(span['size'])
                        except Exception:
                            pass

                        # color
                        rgb = int_to_rgb(span.get('color'))
                        if rgb != (0, 0, 0):
                            try:
                                run.font.color.rgb = RGBColor(*rgb)
                            except Exception:
                                pass

                        # styles
                        flags = span.get('flags', 0)
                        if flags & 16:
                            run.bold = True
                        if flags & 2:
                            run.italic = True
                        if flags & 1:
                            run.font.superscript = True

                    # --- determine if we should insert a visible space between this span and the next ---
                    need_space = False
                    if si + 1 < len(spans):
                        next_span = spans[si + 1]
                        try:
                            # gap in PDF points between current span right and next span left
                            gap_pts = next_span['bbox'][0] - span['bbox'][2]
                        except Exception:
                            gap_pts = 0.0

                        # Use font size (in pts) as a scale for what counts as a visible gap
                        font_pt = span.get('size') or 12.0
                        # threshold: small minimum (1pt) plus a fraction of font size
                        gap_threshold = max(1.0, 0.18 * font_pt)

                        # Next text starts with punctuation? then usually NO space even if gap exists
                        next_text = (next_span.get('text') or '').lstrip()
                        if gap_pts > gap_threshold and not re.match(r'^[,.;:?!\)\]\}%/]', next_text):
                            need_space = True

                    if need_space:
                        # insert a space run with the same font properties so Word renders it consistently
                        space_run = p.add_run(' ')
                        try:
                            space_run.font.name = map_font_name(span.get('font'))
                            if span.get('size'):
                                space_run.font.size = Pt(span.get('size'))
                            if rgb != (0, 0, 0):
                                try:
                                    space_run.font.color.rgb = RGBColor(*rgb)
                                except Exception:
                                    pass
                            if flags & 16:
                                space_run.bold = True
                            if flags & 2:
                                space_run.italic = True
                        except Exception:
                            pass


# ==========================================
# Main
# ==========================================

def main():
    parser = argparse.ArgumentParser(description="High-Fidelity PDF to DOCX")
    parser.add_argument("input", help="Input PDF file")
    parser.add_argument("-o", "--output", help="Output DOCX file", default=None)
    args = parser.parse_args()

    input_file = args.input
    if not args.output:
        args.output = os.path.splitext(input_file)[0] + ".docx"

    if not os.path.exists(input_file):
        print(f"Error: File {input_file} not found.")
        sys.exit(1)

    print(f"Processing: {input_file}")
    
    doc_pdf = fitz.open(input_file)
    doc_word = Document()
    
    # Remove default Section to replace with PDF-sized sections
    # (Actually we just modify the existing first section)
    
    total_pages = len(doc_pdf)

    for i in range(total_pages):
        print(f"  - Converting Page {i+1}/{total_pages}...")
        page = doc_pdf[i]
        
        # Extract
        elements = extract_page_content(page)
        
        # Write
        # For multi-page, we need to ensure sections handle the breaks
        # If it's not the first page, we might need a new section for size changes
        # But for simplicity in this script, we assume uniform page sizes or just break.

        if i > 0:
            # Create a new section for this page to isolate layout/margins/size
            doc_word.add_section()
        
        # Pass the specific section index or let write_to_docx handle the last section

        write_to_docx(doc_word, elements, page.rect.width, page.rect.height)

    doc_word.save(args.output)
    print(f"Success! Saved to {args.output}")

if __name__ == "__main__":
    main()
