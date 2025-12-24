import io
import math
from typing import Dict, Tuple
from reportlab.lib.utils import ImageReader
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

# ==========================================
# Watermark Renderer
# ==========================================

class WatermarkRenderer:
    """
    Handles the generation of watermark PDF pages using ReportLab.
    
    This class is responsible for:
    1. Creating an in-memory PDF stream for the watermark.
    2. Calculating geometry (position, rotation, scaling).
    3. Drawing text or images onto the canvas.
    4. Caching generated pages to optimize performance for uniform page sizes.
    """

    def __init__(self, config: WatermarkConfig):
        self.config = config
        # Cache key: (width, height), Value: pypdf.PageObject
        self._cache: Dict[Tuple[float, float], PageObject] = {}

    def get_watermark(self, page_width: float, page_height: float) -> PageObject:
        """
        Retrieves a watermark PageObject for the specified dimensions.
        Returns a cached object if available, otherwise renders a new one.
        """
        # Round dimensions to avoid cache misses on negligible float differences
        key = (round(page_width, 2), round(page_height, 2))
        
        if key not in self._cache:
            self._cache[key] = self._render_watermark_page(page_width, page_height)
            
        return self._cache[key]

    def _render_watermark_page(self, width: float, height: float) -> PageObject:
        """Internal method to draw the watermark on a fresh PDF page."""
        packet = io.BytesIO()
        
        # Create a new PDF canvas with the specific page size
        c = canvas.Canvas(packet, pagesize=(width, height))
        
        # Apply Global Opacity
        # Note: ReportLab handles alpha via fillAlpha/strokeAlpha
        c.setFillAlpha(self.config.opacity)
        c.setStrokeAlpha(self.config.opacity)

        if self.config.watermark_type == WatermarkType.TEXT:
            self._draw_text(c, width, height)
        elif self.config.watermark_type == WatermarkType.IMAGE:
            self._draw_image(c, width, height)

        c.save()
        packet.seek(0)
        
        # Create a pypdf PageObject from the generated stream
        reader = PdfReader(packet)
        return reader.pages[0]

    def _draw_text(self, c: canvas.Canvas, page_w: float, page_h: float):
        """Draws text watermark with rotation and positioning."""
        text = self.config.text
        c.setFont(self.config.font_name, self.config.font_size)
        c.setFillColor(self.config.font_color)

        # Calculate text width using ReportLab's stringWidth
        text_w = c.stringWidth(text, self.config.font_name, self.config.font_size)
        # Approximate height (ascent) for positioning logic
        text_h = self.config.font_size 

        # Determine coordinates based on alignment
        x, y = self._calculate_position(text_w, text_h, page_w, page_h)

        # Apply Rotation logic
        # To rotate around the center of the text, we translate, rotate, translate back
        c.saveState()
        c.translate(x + text_w / 2, y + text_h / 2)  # Move to center of text
        c.rotate(self.config.rotation)
        c.drawCentredString(0, -text_h / 4, text) # Draw relative to new origin
        c.restoreState()

    def _draw_image(self, c: canvas.Canvas, page_w: float, page_h: float):
        """Draws image watermark with scaling and positioning."""
        img_path = str(self.config.image_path)
        
        # Use ImageReader to get dimensions without loading full image into canvas yet
        utils_img = ImageReader(img_path)
        img_orig_w, img_orig_h = utils_img.getSize()

        # Calculate target dimensions
        target_w = img_orig_w * self.config.image_scale
        target_h = img_orig_h * self.config.image_scale

        # Optional: Scale relative to page width (overrides fixed scale)
        if self.config.image_rel_page > 0:
            target_w = page_w * self.config.image_rel_page
            aspect = img_orig_h / img_orig_w
            target_h = target_w * aspect

        # Determine coordinates
        x, y = self._calculate_position(target_w, target_h, page_w, page_h)

        # Apply Rotation
        c.saveState()
        c.translate(x + target_w / 2, y + target_h / 2)
        c.rotate(self.config.rotation)
        # Draw image centered on the translated origin
        # Note: (x, y, w, h) -> draw relative to -w/2, -h/2
        c.drawImage(
            img_path, 
            -target_w / 2, 
            -target_h / 2, 
            width=target_w, 
            height=target_h, 
            mask='auto',
            preserveAspectRatio=True
        )
        c.restoreState()

    def _calculate_position(self, item_w: float, item_h: float, page_w: float, page_h: float) -> Tuple[float, float]:
        """
        Calculates the bottom-left (x, y) coordinates for the item 
        based on the configured position and margins.
        """
        pos = self.config.position
        mx = self.config.margin_x
        my = self.config.margin_y

        if pos == WatermarkPosition.CENTER:
            return (page_w - item_w) / 2, (page_h - item_h) / 2
        
        elif pos == WatermarkPosition.TOP_LEFT:
            return mx, page_h - item_h - my
            
        elif pos == WatermarkPosition.TOP_RIGHT:
            return page_w - item_w - mx, page_h - item_h - my
            
        elif pos == WatermarkPosition.BOTTOM_LEFT:
            return mx, my
            
        elif pos == WatermarkPosition.BOTTOM_RIGHT:
            return page_w - item_w - mx, my
            
        return 0, 0 # Fallback
