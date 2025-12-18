import os
import re
import base64
import posixpath
from ebooklib import epub
from bs4 import BeautifulSoup, NavigableString
from weasyprint import HTML, CSS
# New dependency for image dimension extraction
try:
    from PIL import Image
    import io
except ImportError:
    # Print warning only once
    if 'Image' not in globals() or Image is not None:
        print("WARNING: Pillow (PIL) is not installed. Image dimension inference will be disabled.")
        print("Install with: pip install Pillow")
    Image = None
    io = None


# --- Helper Functions ---

# Helper: normalize posix path inside epub
def _resolve_href(base_path, href):
    if not href:
        return None
    # Remove leading fragment (e.g., #anchor) before joining, but keep it for later if it exists
    path_only = href.split('#')[0]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_path), path_only))

# Embed resources referenced via url(...) in CSS with data: URIs (images/fonts/etc.)
def _embed_resources_in_css(css_text, css_item_path, items_by_path):
    # preserve local(...) tokens and absolute/external urls
    def repl_url(match):
        raw = match.group(1).strip().strip("\"'")
        if not raw or raw.lower().startswith('data:') or re.match(r'^[a-zA-Z]+://', raw):
            return f'url("{raw}")'  # already fine
        
        # Resolve path relative to the CSS file
        resolved = _resolve_href(css_item_path, raw)
        resource = items_by_path.get(resolved)
        
        if not resource:
            # fallback: keep original
            return f'url("{raw}")'
        
        # embed as base64
        content = resource.get_content()
        try:
            b64 = base64.b64encode(content).decode('ascii')
        except Exception:
            # as fallback, attempt latin-1 decode then re-encode (rare)
            b64 = base64.b64encode(content).decode('ascii', errors='ignore')
        media_type = getattr(resource, 'media_type', None) or 'application/octet-stream'
        return f'url("data:{media_type};base64,{b64}")'

    # Replace url(...) occurrences
    css_text = re.sub(r'url\(([^)]+)\)', repl_url, css_text, flags=re.IGNORECASE)

    # Some EPUB CSS uses @import "..." rules; try to inline them too if possible
    def repl_import(m):
        # Group 1 is the quoted URL from the @import rule
        href = m.group(1).strip().strip("\"'")
        resolved = _resolve_href(css_item_path, href)
        resource = items_by_path.get(resolved)
        if resource:
            try:
                css_inner = resource.get_content().decode('utf-8')
            except Exception:
                css_inner = resource.get_content().decode('latin-1', errors='ignore')
            css_inner = _embed_resources_in_css(css_inner, resource.file_name, items_by_path)
            # Replace the @import rule with the content and a comment for debugging
            return f"/* Inlined @import: {href} */\n{css_inner}\n"
        return m.group(0)  # leave as-is if not found

    css_text = re.sub(r'@import\s+["\']([^"\']+)["\']\s*;', repl_import, css_text, flags=re.IGNORECASE)

    return css_text

# NEW LOGIC: Extract pixel dimensions from an image resource
def _get_image_dimensions(resource):
    """
    Returns (width, height) in pixels, or (None, None) if dimensions cannot be determined.
    Requires the PIL/Pillow library.
    """
    if Image is None or io is None:
        return None, None

    content = resource.get_content()
    try:
        img = Image.open(io.BytesIO(content))
        return img.size
    except Exception:
        return None, None


# Inline images, links, audio/video, and other src attributes present in XHTML
def _process_and_embed_resources_in_soup(soup, base_path, items_by_path, debug_log=None):
    
    # img src processing
    for img in soup.find_all('img'):
        src = img.get('src')
        if not src:
            continue
            
        resolved = _resolve_href(base_path, src)
        res_item = items_by_path.get(resolved)
        
        if res_item:
            b64 = base64.b64encode(res_item.get_content()).decode('ascii')
            media_type = res_item.media_type or 'image/png'
            
            # --- CORRECTED Image Logic ---
            # 1. Get existing style to preserve author's special effects (like borders)
            style_attr = img.get('style', '')

            # 2. Define the safeguard style
            # max-width: 100% -> Ensures image never exceeds the page width (Fixes "too large/out of page")
            # height: auto    -> Preserves aspect ratio so image doesn't stretch
            # display: block  -> Required to make 'margin: auto' work for centering
            # margin: 0 auto  -> Centers the image horizontally (Fixes "aligned to right")
            safeguard_style = "max-width: 100% !important; height: auto !important; display: block; margin: 0 auto;"
            
            # 3. Apply the safeguard ONLY if the image doesn't already have these properties explicitly set.
            # We prepend our safeguard so it acts as the default base.
            if "max-width" not in style_attr:
                img['style'] = f"{safeguard_style} {style_attr}".strip()
                if debug_log is not None:
                    debug_log.append(f"FIX_IMG: Applied centering and max-width to {src}")

            # Set data URI
            img['src'] = f'data:{media_type};base64,{b64}'
        else:
            # try to keep external URLs intact; if relative and missing, remove to avoid broken inline
            if not re.match(r'^[a-zA-Z]+://', src) and not src.startswith('data:'):
                img.decompose()
                if debug_log is not None:
                    debug_log.append(f"REMOVE_IMG: Decomposed missing relative image: {src}")

    # srcset attribute (could be a comma-separated list)
    for tag in soup.find_all(attrs={'srcset': True}):
        srcset = tag['srcset']
        parts = []
        for piece in srcset.split(','):
            piece = piece.strip()
            if not piece:
                continue
            url_part = piece.split()[0]
            tail = ' '.join(piece.split()[1:])  # like "2x" or "800w"
            resolved = _resolve_href(base_path, url_part)
            res_item = items_by_path.get(resolved)
            if res_item:
                b64 = base64.b64encode(res_item.get_content()).decode('ascii')
                media_type = res_item.media_type or 'image/png'
                parts.append(f'data:{media_type};base64,{b64} {tail}'.strip())
            else:
                parts.append(piece)  # leave unchanged (external or missing)
        tag['srcset'] = ', '.join(parts)

    # other tags with src (audio, video, source, etc.)
    for tag in soup.find_all(src=True):
        if tag.name == 'img':
            continue
        src = tag['src']
        if not src:
            continue
        resolved = _resolve_href(base_path, src)
        res_item = items_by_path.get(resolved)
        if res_item:
            b64 = base64.b64encode(res_item.get_content()).decode('ascii')
            media_type = res_item.media_type or 'application/octet-stream'
            tag['src'] = f'data:{media_type};base64,{b64}'
        else:
            pass

    # FIX 3: Rewrite internal <a> links
    for tag in soup.find_all('a', href=True):
        href = tag['href']
        if not href or href.startswith('http') or href.startswith('data:') or href.startswith('mailto:'):
            continue
        if href.startswith('#'):
            continue

        resolved_path_with_fragment = posixpath.normpath(posixpath.join(posixpath.dirname(base_path), href))
        path_part, _, fragment_part = resolved_path_with_fragment.partition('#')
        doc_id = f"epub-doc-{path_part.replace('/', '-').replace('.', '_').replace('%', '_')}"
        
        if fragment_part:
            tag['href'] = f"#{fragment_part}"
        else:
            tag['href'] = f"#{doc_id}"
            
    # Remove metadata tags
    tags_to_remove = ['title', 'meta', 'script']
    for tag in soup.find_all(tags_to_remove, recursive=True):
        tag.decompose()
    
    # Remove non-stylesheet links
    for tag in list(soup.find_all('link')):
        if not (tag.get('rel') and 'stylesheet' in tag.get('rel').lower()):
            tag.decompose()


def epub_to_pdf_preserve_style(epub_path, output_pdf, page_size='A4', margin_in_inches=1, debug=False):
    book = epub.read_epub(epub_path)
    debug_log = [] if debug else None

    # Build map: file path -> item
    items_by_path = {getattr(it, 'file_name', ''): it for it in book.get_items() if getattr(it, 'file_name', None)}

    # collect global CSS items in spine/manifest order (so cascade is respected)
    css_items = []
    for it in book.get_items():
        mt = getattr(it, 'media_type', '') or ''
        if mt.lower() in ('text/css',) or (getattr(it, 'file_name', '').lower().endswith('.css')):
            css_items.append(it)

    # Prepare merged global CSS (inlined with embedded resources)
    merged_global_css = []
    for css_item in css_items:
        try:
            css_text = css_item.get_content().decode('utf-8')
        except Exception:
            css_text = css_item.get_content().decode('latin-1', errors='ignore')
        
        # Embed resources (fonts/images) referenced in this CSS
        css_text = _embed_resources_in_css(css_text, css_item.file_name, items_by_path)
        merged_global_css.append(f"/* == Global CSS: {css_item.file_name} == */\n{css_text}\n")
        if debug_log is not None:
             debug_log.append(f"INLINE_CSS: Inlined global CSS: {css_item.file_name}")

    # Build reading order spine
    # Use book.spine and fallback to all xhtml docs
    spine = [entry[0] for entry in getattr(book, 'spine', []) if entry and entry[0] != 'nav']
    spine_items = []
    if not spine:
        # fallback to all xhtml docs
        for it in book.get_items():
            if getattr(it, 'media_type', '') == 'application/xhtml+xml':
                spine_items.append(it)
        if debug_log is not None:
             debug_log.append("SPINE: Used manifest fallback for reading order.")
    else:
        for idref in spine:
            try:
                it = book.get_item_with_id(idref)
            except Exception:
                it = None
            if it and getattr(it, 'media_type', '') == 'application/xhtml+xml':
                spine_items.append(it)
        if debug_log is not None:
             debug_log.append("SPINE: Used explicit EPUB spine for reading order.")


    chapters_html = []
    # Collect document-specific CSS items to ensure correct cascade *after* global CSS
    document_head_styles = []

    for doc_item in spine_items:
        raw = doc_item.get_content()
        # decode (xhtml usually utf-8)
        try:
            raw_text = raw.decode('utf-8')
        except Exception:
            raw_text = raw.decode('latin-1', errors='ignore')

        soup = BeautifulSoup(raw_text, 'lxml')
        
        # Create a unique, linkable ID for this document (used by FIX 3)
        doc_id = f"epub-doc-{doc_item.file_name.replace('/', '-').replace('.', '_').replace('%', '_')}"


        # 1) Inline linked styles (<link rel="stylesheet">) that reference internal css items.
        for link in list(soup.find_all('link', rel=lambda v: v and 'stylesheet' in v.lower())):
            href = link.get('href')
            if not href:
                link.decompose()
                continue
            resolved = _resolve_href(doc_item.file_name, href)
            css_item = items_by_path.get(resolved)
            if css_item:
                try:
                    css_text = css_item.get_content().decode('utf-8')
                except Exception:
                    css_text = css_item.get_content().decode('latin-1', errors='ignore')
                css_text = _embed_resources_in_css(css_text, css_item.file_name, items_by_path)
                
                # Collect to place in the final head (before document body content)
                document_head_styles.append(f"/* == Chapter Link CSS: {doc_item.file_name} -> {resolved} == */\n{css_text}\n")
                if debug_log is not None:
                     debug_log.append(f"INLINE_LINK_CSS: Inlined chapter linked CSS: {href}")
                link.decompose()
            else:
                # Remove external/missing links to prevent broken relative references
                link.decompose()
                if debug_log is not None:
                     debug_log.append(f"REMOVE_LINK_CSS: Removed external/missing link: {href}")

        # 2) Process and collect <style> blocks in doc head/body
        for style in list(soup.find_all('style')): # Use list() so decompose() doesn't mess up iteration
            if style.string and style.string.strip():
                css_text = style.string
                css_text = _embed_resources_in_css(css_text, doc_item.file_name, items_by_path)
                document_head_styles.append(f"/* == Chapter Style Block: {doc_item.file_name} == */\n{css_text}\n")
                if debug_log is not None:
                     debug_log.append(f"INLINE_STYLE_BLOCK: Inlined chapter <style> block.")
            style.decompose()  # we'll reinsert into final head

        # 3) Embed images (img, srcset, etc.), fix internal <a> links, and infer image dimensions
        _process_and_embed_resources_in_soup(soup, doc_item.file_name, items_by_path, debug_log)

        # 4) Extract and modify the body content.
        # Preserve <body> attributes (like id/class) by renaming to <div>
        body = soup.body
        if body:
            # Transfer attributes
            attributes = dict(body.attrs)
            
            # Create the new <div> tag, preserving attributes
            new_div = soup.new_tag('div', **attributes)
            new_div.name = 'div' # Explicitly set name
            
            # Move children from <body> to new <div>
            for child in list(body.contents):
                new_div.append(child)
            
            # Add our own class but preserve existing ones
            existing_classes = new_div.get('class', [])
            new_div['class'] = ['epub-chapter'] + existing_classes
            
            # Add the unique ID for linking (FIX 3)
            new_div['id'] = doc_id
            
            content_html = str(new_div) # Use the whole modified <div>
        else:
            # fallback: use whole document if <body> missing
            content_html = f'<div class="epub-chapter epub-fallback-body" id="{doc_id}">{str(soup.html if soup.html else soup)}</div>'
            if debug_log is not None:
                debug_log.append("FALLBACK_BODY: No <body> found, using fallback <div>.")


        # We append the whole <div>, not just its contents wrapped in a generic <section>
        chapters_html.append(content_html)
        if debug_log is not None:
             debug_log.append(f"CHAPTER_PROCESSED: {doc_item.file_name}")

    # --- Compose final HTML ---
    # Cascade order: WeasyPrint Defaults < Page/Reset CSS < Global EPUB CSS < Chapter-specific CSS
    all_css = '\n'.join(merged_global_css + document_head_styles)

    # Add @page and margin rules + a reset for common body-level styles
    # NOTE: We removed the aggressive body/html resets to preserve author's margins/padding on <body>.
    page_css = f"""
@page {{
    size: {page_size}; 
    margin: {margin_in_inches}in !important; 
    padding: 0 !important;
}} 

/* Ensure chapters are separated by page breaks and avoid awkward breaks mid-chapter */
.epub-chapter {{ 
    page-break-after: always; 
    /* page-break-inside: auto allows the book's own rules to control breaks */
    /* page-break-inside: avoid; */ 
}}
/* Final chapter does not need a break after */
.epub-chapter:last-child {{
    page-break-after: auto;
}}
/* Handle the case where the doc had no body */
.epub-fallback-body {{ 
    page-break-after: always; 
}}
"""

    final_html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{getattr(book, 'title', 'EPUB to PDF')}</title>
<style>
/* == PAGE & RESET CSS (Lowest Precedence) == */
{page_css}
/* == EPUB STYLES START (Highest Precedence) == */
{all_css}
/* == EPUB STYLES END == */
</style>
</head>
<body>
{''.join(chapters_html)}
</body>
</html>
"""
    if debug_log is not None:
        print("\n--- DEBUG LOG ---")
        for log in debug_log:
            print(log)
        print("-----------------\n")

    # Render to PDF
    try:
        # WeasyPrint's HTML parser can take the string directly.
        # The internal styles will be applied to the document content.
        HTML(string=final_html).write_pdf(output_pdf)
        print(f"✅ Saved PDF: {output_pdf}")
    except Exception as e:
        print(f"❌ Error during PDF rendering (WeasyPrint): {e}")
        # Suggest a fallback if rendering fails, though it's usually a dependency issue
        print("SUGGESTION: Check WeasyPrint dependencies (pango, cairo, etc.)")


def convert_all_epubs_in_directory(directory=".", debug=False):
    """
    Convert every EPUB file in the given directory to PDF.
    """
    epub_files = [f for f in os.listdir(directory) if f.lower().endswith(".epub")]

    if not epub_files:
        print("⚠️ No EPUB files found in this directory.")
        return

    print(f"📚 Found {len(epub_files)} EPUB file(s) in '{os.path.abspath(directory)}':\n")
    
    converted_files = 0
    for epub_file in epub_files:
        epub_path = os.path.join(directory, epub_file)
        pdf_output = os.path.splitext(epub_path)[0] + ".pdf"
        print(f"--- Converting: {epub_file} ---")
        try:
            # Check if Pillow is missing before starting
            if Image is None and ("INFER_IMG" in epub_to_pdf_preserve_style.__code__.co_names or "_get_image_dimensions" in epub_to_pdf_preserve_style.__code__.co_names):
                print("NOTE: Image dimension inference is disabled (Pillow not installed). Images may be incorrectly scaled.")

            epub_to_pdf_preserve_style(epub_path, pdf_output, debug=debug) # Pass debug flag
            converted_files += 1
        except Exception as e:
            print(f"❌ Error converting {epub_file}: {e}")
            # Optionally, re-raise if you want to stop on error
            # raise e

    print(f"\n🎉 All conversions completed ({converted_files}/{len(epub_files)} successful)!")


# --- Main execution ---
#
# Added an optional `debug=True` to the main function call.
#
# --- OPTION 1: Convert all EPUBs in the current directory (Default) ---
# NOTE: Set debug=True to enable the debug log during conversion.
DEBUG_MODE = False
print(f"--- Starting bulk conversion (Debug Mode: {'ON' if DEBUG_MODE else 'OFF'}) for all .epub files in this directory ---")
convert_all_epubs_in_directory(".", debug=DEBUG_MODE)

# --- OPTION 2: Example for converting a single file ---
#
# print("\n--- Starting single file conversion ---")
# epub_file = "example.epub"      # Path to your EPUB file
# pdf_output = "output.pdf"       # Output PDF file name
#
# if os.path.exists(epub_file):
#     try:
#         # Use the modified function and pass debug mode
#         epub_to_pdf_preserve_style(epub_file, pdf_output, debug=True)
#     except Exception as e:
#         print(f"❌ Error converting {epub_file}: {e}")
# else:
#     print(f"❌ Single file not found: {epub_file}. Please check the file path.")
