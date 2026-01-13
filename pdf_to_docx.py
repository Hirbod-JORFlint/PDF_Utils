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

def map_font_name(fname, flags):
    """
    Map PDF font names to standard Word fonts to improve visual fidelity.
    """
	if not fname: return "Calibri"
		
		# Remove CID or Identity-H suffixes
	fname = fname.split('+')[-1] # Remove subset prefix
	fname = fname.split('-')[0]  # Remove encoding suffixes
    
    fname_lower = fname.lower()
    # strip 6-letter subset prefix if present
    fname_clean = re.sub(r'^[A-Z]{6}\+', '', fname)
    # remove common suffixes/tokens and punctuation that often appear in PDF font names
    fname_clean = re.sub(r'[-_,.](bold|italic|regular|mt|ps|std|roman|oblique|semi|condensed|narrow)$', '', fname_clean, flags=re.I)
    fname_clean = fname_clean.strip()
    fname_lower = fname_clean.lower()

    # Common family mappings
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

    # Final fallback: return cleaned name or Calibri if empty
    return fname_clean if fname_clean else "Calibri"

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
    """
    Extracts content including Hyperlinks and Layout details.
    """
    elements = []
    
    # 0. Get Links
    links = page.get_links()

    # 1. Detect Tables
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

    # 2. Extract Text & Images
    text_blocks = page.get_text("dict")["blocks"]
    
    for b in text_blocks:
        bbox = b['bbox']
        
        # Skip if inside table
        if any(is_box_inside(bbox, tr) for tr in table_rects):
            continue

        if b['type'] == 0:  # Text
            para_lines = []
            for line in b['lines']:
                line_spans = []
                for span in line['spans']:
                    # Check for link
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
				# Merge consecutive PDF lines into paragraphs using a vertical-gap heuristic.
				paragraphs = []
				def avg_font_size(line):
					sizes = [s.get('size') for s in line['spans'] if s.get('size')]
					return float(sizes[0]) if sizes else 12.0

				# Start with first line as current paragraph
				current = {'bbox': para_lines[0]['bbox'], 'spans': list(para_lines[0]['spans'])}
				prev_center_y = (para_lines[0]['bbox'][1] + para_lines[0]['bbox'][3]) / 2.0
				current_font = avg_font_size(para_lines[0])

				for ln in para_lines[1:]:
					center_y = (ln['bbox'][1] + ln['bbox'][3]) / 2.0
					gap = abs(center_y - prev_center_y)

					# threshold = ~0.6–1.0 * font-size (empirical); adjust if needed
					threshold = max(6.0, 0.75 * current_font)

					if gap <= threshold:
                        # Check for hyphenation at the end of the previous line
                        last_text = current['spans'][-1]['text'].strip()
                        if last_text.endswith('-') and len(last_text) > 3:
                            # Remove hyphen and don't add space
                            current['spans'][-1]['text'] = current['spans'][-1]['text'].rstrip('-')
                        else:
                            # Add a space between lines if the last span doesn't end in one
                            if not current['spans'][-1]['text'].endswith(' '):
                                current['spans'][-1]['text'] += ' '
                        
                        current['spans'].extend(ln['spans'])
                        prev_center_y = center_y
					else:
						# Close current paragraph and start a new one
						paragraphs.append(current)
						current = {'bbox': ln['bbox'], 'spans': list(ln['spans'])}
						prev_center_y = center_y
						current_font = avg_font_size(ln)

				paragraphs.append(current)

				# Add one element per logical paragraph (each element.lines is a list of 1 item: the merged paragraph)
				for para in paragraphs:
					elements.append({
						'type': 'text',
						'bbox': bbox,
						'y_sort': para['bbox'][1] if isinstance(para['bbox'], (list, tuple)) else bbox[1],
						'lines': [para],
						'indent_pt': bbox[0]
					})
					
        elif b['type'] == 1:  # Image
            img_bytes = b.get("image")
            ext = b.get("ext", "png")
            if img_bytes:
                elements.append({
                    'type': 'image',
                    'bbox': bbox,
                    'y_sort': bbox[1],
                    'bytes': img_bytes,
                    'ext': ext
                })

    # Sort by Vertical position
    elements.sort(key=lambda x: x['y_sort'])
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
            
            for r, row_data in enumerate(data):
                for c, cell_text in enumerate(row_data):
                    if c < len(table.rows[r].cells):
                        cell = table.rows[r].cells[c]
                        # Clean None values
                        t_val = cell_text if cell_text else ""
                        cell.text = t_val.strip()
            
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
            pf = p.paragraph_format
            
            # 1. Indentation
            # Calculate relative indent from the left margin buffer
            raw_indent = el['indent_pt']
            relative_indent = raw_indent - margin_buffer
            if relative_indent > 0:
                pf.left_indent = Pt(relative_indent)
            
            # 2. Spacing (Tight packing)
            # Remove space after paragraph to prevent giant gaps
            pf.space_after = Pt(0)
            
            # 3. Process Runs
            for line in el['lines']:
                for span in line['spans']:
                    text = span['text']
                    
                    # Handle Hyperlinks
                    if span['link']:
                        rgb = int_to_rgb(span['color'])
                        is_bold = bool(span['flags'] & 16)
                        is_italic = bool(span['flags'] & 2)
                        fname = map_font_name(span['font'], span['flags'])
                        add_hyperlink(p, span['link'], text, rgb, is_bold, is_italic, fname, span['size'])
                        continue

                    if not text: continue
                    
                    run = p.add_run(text)
                    
                    # Font Name
                    run.font.name = map_font_name(span['font'], span['flags'])
                    
                    # Size
                    try:
                        run.font.size = Pt(span['size'])
                    except:
                        pass
                    
                    # Color
                    rgb = int_to_rgb(span['color'])
                    if rgb != (0, 0, 0):
                        run.font.color.rgb = RGBColor(*rgb)
                    
                    # Styles based on flags
                    # bit 0: superscript
                    # bit 1: italic
                    # bit 2: serif
                    # bit 3: monospaced
                    # bit 4: bold
                    flags = span['flags']
                    if flags & 16:
                        run.bold = True
                    if flags & 2:
                        run.italic = True
                    if flags & 1:
                        run.font.superscript = True
                        
                # Add implicit space or break?
				# preserve intra-span spacing: add a space at the end only when needed
				txt = text
				if txt and not txt.endswith(' '):
					# peek next span text if available (to decide whether a space is needed)
					# If there's no easy peek (complexity), it's safe to append a single space here
					# to avoid word concatenation — the merge step above already reduces line breaks.
					txt = txt
				run = p.add_run(txt)

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
            doc_word.add_page_break()
            
        write_to_docx(doc_word, elements, page.rect.width, page.rect.height)

    doc_word.save(args.output)
    print(f"Success! Saved to {args.output}")

if __name__ == "__main__":
    main()
