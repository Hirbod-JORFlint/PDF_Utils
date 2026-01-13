#!/usr/bin/env python3
"""
pdf_to_text.py

Convert PDF -> plain text (.txt) preserving layout heuristics:
 - Attempts to reconstruct paragraph runs and spacing between spans.
 - Detects links and appends inline URL markers (optional).
 - Emits table rows as tab-separated lines.
 - Optionally inserts image placeholders.

Usage:
  python pdf_to_text.py input.pdf
  python pdf_to_text.py input.pdf -o out.txt --preserve-links --include-images
"""

import argparse
import os
import sys
import re
import io

import fitz  # PyMuPDF

# -------------------------
# Helpers (adapted)
# -------------------------
def is_box_inside(inner, outer):
    ix0, iy0, ix1, iy1 = inner
    ox0, oy0, ox1, oy1 = outer
    cx = (ix0 + ix1) / 2
    cy = (iy0 + iy1) / 2
    return (ox0 <= cx <= ox1) and (oy0 <= cy <= oy1)

def get_link_target(bbox, links):
    """Find if a bbox intersects with a known link area (returns uri or None)."""
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    
    for link in links:
        lb = link.get('from')  # Rect
        if lb is None:
            continue
        if lb.x0 <= cx <= lb.x1 and lb.y0 <= cy <= lb.y1:
            return link.get('uri')
    return None

def extract_page_content(page):
    """
    Extract page content into a list of elements like:
      {'type': 'text', 'bbox':..., 'y_sort':..., 'lines': [...], 'indent_pt': ...}
      {'type': 'table', 'bbox':..., 'y_sort':..., 'data': [...]}
      {'type': 'image', 'bbox':..., 'y_sort':..., 'bytes':..., 'ext': ...}
    Uses similar paragraph merging heuristics as original script.
    """
    elements = []
    links = page.get_links()

    # Tables
    table_rects = []
    try:
        tables = page.find_tables()
    except Exception:
        tables = []
    for tab in tables:
        bbox = tab.bbox
        table_rects.append(bbox)
        elements.append({
            'type': 'table',
            'bbox': bbox,
            'y_sort': bbox[1],
            'data': tab.extract()
        })

    # Text & Images
    try:
        text_blocks = page.get_text("dict")["blocks"]
    except Exception:
        text_blocks = []

    for b in text_blocks:
        bbox = b.get('bbox', (0,0,0,0))
        if any(is_box_inside(bbox, tr) for tr in table_rects):
            continue

        if b.get('type') == 0:  # Text
            para_lines = []
            for line in b.get('lines', []):
                line_spans = []
                for span in line.get('spans', []):
                    uri = get_link_target(span.get('bbox', (0,0,0,0)), links)
                    line_spans.append({
                        'text': span.get('text', ''),
                        'font': span.get('font'),
                        'size': span.get('size'),
                        'color': span.get('color'),
                        'flags': span.get('flags', 0),
                        'bbox': span.get('bbox'),
                        'origin': span.get('origin'),
                        'link': uri
                    })
                if line_spans:
                    para_lines.append({'bbox': line.get('bbox'), 'spans': line_spans})

            if para_lines:
                # Merge lines into paragraphs using vertical gap heuristic
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
                    min_gap_pts = 3.0
                    rel_gap_factor = 0.45
                    threshold = max(min_gap_pts, rel_gap_factor * current_font)

                    if gap <= threshold:
                        # join
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

        elif b.get('type') == 1:  # Image
            img_bytes = b.get("image")
            if img_bytes:
                elements.append({
                    'type': 'image',
                    'bbox': bbox,
                    'y_sort': bbox[1],
                    'bytes': img_bytes,
                    'ext': b.get("ext", "png")
                })

    elements.sort(key=lambda x: x['y_sort'])
    return elements

# -------------------------
# Text writing
# -------------------------
def stringify_elements(elements, page_num, page_width=None, preserve_links=True, include_images=False):
    """Return a list of text lines representing the elements for a page."""
    out_lines = []

    for el in elements:
        if el['type'] == 'table':
            data = el.get('data') or []
            for row in data:
                # convert row into tab-separated string, cleaning None -> ""
                out_lines.append("\t".join([(cell or "").strip() for cell in row]))
            out_lines.append("")  # blank line after table

        elif el['type'] == 'image':
            if include_images:
                bbox = el.get('bbox')
                out_lines.append(f"[IMAGE on page {page_num} bbox={bbox}]")
                out_lines.append("")  # blank line

        elif el['type'] == 'text':
            # create one paragraph string by concatenating spans, deciding spaces by bbox gaps
            para_text = ""
            lines = el.get('lines', [])
            for line in lines:
                spans = line.get('spans', [])
                for si, span in enumerate(spans):
                    text = span.get('text') or ''
                    if not text:
                        continue

                    if preserve_links and span.get('link'):
                        # inline the URL marker
                        uri = span.get('link')
                        # prefer compact representation: "label [url]"
                        para_text += text
                        if uri:
                            # ensure single space before link marker if needed
                            if not para_text.endswith(' '):
                                para_text += ' '
                            para_text += f"[{uri}]"
                    else:
                        para_text += text

                    # spacing heuristic between spans
                    need_space = False
                    if si + 1 < len(spans):
                        next_span = spans[si + 1]
                        try:
                            gap_pts = next_span['bbox'][0] - span['bbox'][2]
                        except Exception:
                            gap_pts = 0.0

                        font_pt = span.get('size') or 12.0
                        gap_threshold = max(1.0, 0.18 * font_pt)

                        next_text = (next_span.get('text') or '').lstrip()
                        if gap_pts > gap_threshold and not re.match(r'^[,.;:?!\)\]\}%/]', next_text):
                            need_space = True

                    if need_space:
                        para_text += ' '

                # if paragraph has multiple 'lines' entries (should be rare here), treat them as continued
            out_lines.append(para_text.rstrip())
    return out_lines

# -------------------------
# Main
# -------------------------
def main():
    parser = argparse.ArgumentParser(description="PDF -> plain text exporter (preserves spacing heuristics).")
    parser.add_argument("input", help="Input PDF file")
    parser.add_argument("-o", "--output", help="Output text file (default: input.txt)", default=None)
    parser.add_argument("--preserve-links", action="store_true", help="Append inline URLs for detected links (text [URL])")
    parser.add_argument("--include-images", action="store_true", help="Insert image placeholders into text output")
    parser.add_argument("--page-sep", choices=['formfeed','line'], default='line', help="Page separator style")
    args = parser.parse_args()

    input_file = args.input
    if not args.output:
        args.output = os.path.splitext(input_file)[0] + ".txt"

    if not os.path.exists(input_file):
        print(f"Error: File {input_file} not found.")
        sys.exit(1)

    doc_pdf = fitz.open(input_file)
    total_pages = len(doc_pdf)

    out_lines_all = []

    for i in range(total_pages):
        page_num = i + 1
        page = doc_pdf[i]
        elements = extract_page_content(page)
        page_lines = stringify_elements(elements, page_num, page_width=page.rect.width,
                                        preserve_links=args.preserve_links, include_images=args.include_images)

        # write header optionally
        out_lines_all.append(f"---- Page {page_num}/{total_pages} ----")
        out_lines_all.extend(page_lines)

        # page separator
        if args.page_sep == 'formfeed':
            out_lines_all.append("\f")
        else:
            out_lines_all.append("\n")  # blank line between pages

    # Save to file (utf-8)
    try:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines_all).rstrip() + "\n")
        print(f"Saved text to {args.output}")
    except Exception as e:
        print(f"Error writing output: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
