import sys
import os
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Optional, Tuple, Union, List

# Check for required libraries at import time
try:
    from reportlab.lib import colors
    from reportlab.lib.colors import Color
    from reportlab.pdfgen import canvas
    from pypdf import PdfReader, PdfWriter, PageObject
except ImportError as e:
    raise ImportError(f"Missing required dependency: {e}. Please install 'reportlab' and 'pypdf'.")

from WatermarkConfig import WatermarkConfig
from WatermarkRenderer import WatermarkRenderer

# ==========================================
# PDF Processor
# ==========================================

class PDFProcessor:
    """
    Manages the workflow of loading, processing, and saving the PDF.
    
    Responsibilities:
    1. File I/O (Safe loading and saving).
    2. Handling PDF Encryption.
    3. Parsing page selection (ranges, indices).
    4. Iterating pages and applying the watermark via WatermarkRenderer.
    """

    def __init__(self, input_path: str, output_path: str, config: WatermarkConfig, password: Optional[str] = None):
        self.input_path = Path(input_path)
        self.output_path = Path(output_path)
        self.config = config
        self.password = password
        self.renderer = WatermarkRenderer(config)
        self.reader: Optional[PdfReader] = None
        self.writer = PdfWriter()

    def load_pdf(self):
        """Loads the PDF and handles decryption if necessary."""
        if not self.input_path.exists():
            raise InvalidInputError(f"Input file not found: {self.input_path}")

        try:
            self.reader = PdfReader(self.input_path)
            
            # Handle Encryption
            if self.reader.is_encrypted:
                if self.password:
                    self.reader.decrypt(self.password)
                else:
                    # Attempt empty password (common for some restricted PDFs)
                    try:
                        self.reader.decrypt("")
                    except:
                        pass
                
                # Check if still locked
                if self.reader.is_encrypted:
                    raise PDFProcessingError("PDF is encrypted. Please provide a valid password.")
                    
        except Exception as e:
            if isinstance(e, PDFProcessingError):
                raise e
            raise PDFProcessingError(f"Failed to load PDF: {e}")

    def process(self, page_selection: Union[str, List[int], None] = None):
        """
        Main execution loop.
        
        Args:
            page_selection: Defines which pages to watermark.
                - None: All pages (default)
                - List[int]: Specific 0-based indices (e.g., [0, 2, 5])
                - str: Range string (e.g., "0, 2-5, 8")
        """
        if self.reader is None:
            self.load_pdf()

        total_pages = len(self.reader.pages)
        target_indices = self._parse_page_selection(page_selection, total_pages)

        print(f"Processing {len(target_indices)} pages out of {total_pages}...")

        try:
            for i, page in enumerate(self.reader.pages):
                # 1. Always add the page to the writer
                # We work on the page object directly before adding it to writer
                
                if i in target_indices:
                    self._apply_watermark_to_page(page)
                
                self.writer.add_page(page)
                
            # Copy over metadata (optional but recommended)
            if self.reader.metadata:
                self.writer.add_metadata(self.reader.metadata)

        except Exception as e:
            raise PDFProcessingError(f"Error during processing loop: {e}")

    def _apply_watermark_to_page(self, page: PageObject):
        """Merges the generated watermark onto a single PDF page."""
        # Get dimensions from the MediaBox (physical page size)
        # float() cast ensures compatibility with reportlab
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        # Get the watermark stamp for these specific dimensions
        watermark_page = self.renderer.get_watermark(page_width, page_height)
        
        # Merge watermark ON TOP of the existing page
        # Note: If you wanted the watermark behind, you'd use merge_page(watermark, over=False)
        # but standard watermarking usually goes on top.
        page.merge_page(watermark_page)

    def save_pdf(self):
        """Writes the result to disk."""
        try:
            # Ensure output directory exists
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.output_path, "wb") as f:
                self.writer.write(f)
            
            print(f"Successfully saved to: {self.output_path}")
            
        except OSError as e:
            raise PDFProcessingError(f"Failed to save output PDF: {e}")

    def _parse_page_selection(self, selection: Union[str, List[int], None], total_pages: int) -> set:
        """
        Parses user input into a set of unique 0-based page indices.
        
        Supports:
        - None -> All pages
        - [0, 2] -> Pages 0 and 2
        - "0, 3-5" -> Pages 0, 3, 4, 5
        """
        if selection is None:
            return set(range(total_pages))

        indices = set()

        if isinstance(selection, list):
            for idx in selection:
                if 0 <= idx < total_pages:
                    indices.add(idx)
                    
        elif isinstance(selection, str):
            # Parse string format "1, 3-5, 7"
            parts = selection.split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    try:
                        start, end = map(int, part.split('-'))
                        # Handle standard human range (inclusive)
                        # Ensure bounds
                        start = max(0, start)
                        end = min(total_pages - 1, end)
                        if start <= end:
                            indices.update(range(start, end + 1))
                    except ValueError:
                        continue # Ignore malformed ranges
                else:
                    try:
                        idx = int(part)
                        if 0 <= idx < total_pages:
                            indices.add(idx)
                    except ValueError:
                        continue

        return indices
