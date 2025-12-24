#!/usr/bin/env python3
"""
PDF Watermark Adder Module - Part 1: Configuration & Infrastructure

This module provides a robust, production-ready solution for adding text or 
image watermarks to existing PDF files.

Architecture:
1. Configuration: Data classes and Enums to define watermark properties.
2. Rendering: ReportLab generation of watermark stamps (in-memory).
3. Processing: pypdf integration to merge stamps with source PDFs.
4. CLI: Command-line interface for standalone usage.

Dependencies:
- reportlab
- pypdf
"""

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

# ==========================================
# Custom Exceptions
# ==========================================

class WatermarkError(Exception):
    """Base exception for all watermarking operations."""
    pass

class InvalidInputError(WatermarkError):
    """Raised when input parameters or files are invalid."""
    pass

class PDFProcessingError(WatermarkError):
    """Raised when the PDF processing/merging fails."""
    pass

class ResourceError(WatermarkError):
    """Raised when external resources (fonts, images) cannot be loaded."""
    pass

# ==========================================
# Enumerations & Constants
# ==========================================

class WatermarkType(Enum):
    """Defines the mode of watermarking."""
    TEXT = "text"
    IMAGE = "image"

class WatermarkPosition(Enum):
    """
    Defines anchor points for watermark placement.
    """
    CENTER = "center"
    TOP_LEFT = "top_left"
    TOP_RIGHT = "top_right"
    BOTTOM_LEFT = "bottom_left"
    BOTTOM_RIGHT = "bottom_right"

# ==========================================
# Configuration Data Class
# ==========================================

@dataclass
class WatermarkConfig:
    """
    Configuration object holding all style and geometry settings for the watermark.
    
    This class handles validation of inputs immediately upon instantiation.
    """
    
    # --- Core Settings ---
    watermark_type: WatermarkType
    
    # --- Geometry & Appearance ---
    position: WatermarkPosition = WatermarkPosition.CENTER
    rotation: float = 0.0          # Degrees (counter-clockwise)
    opacity: float = 0.5           # 0.0 (transparent) to 1.0 (solid)
    margin_x: float = 20.0         # Horizontal margin in points
    margin_y: float = 20.0         # Vertical margin in points
    
    # --- Text Specific Settings ---
    text: Optional[str] = None
    font_name: str = "Helvetica"
    font_size: int = 40
    font_color: Color = colors.grey # ReportLab Color object
    
    # --- Image Specific Settings ---
    image_path: Optional[Union[str, Path]] = None
    image_scale: float = 1.0       # Scale relative to original image size
    image_rel_page: float = 0.0    # If > 0, scale image to occupy this % of page width (0.0-1.0)
    
    def __post_init__(self):
        """Validates configuration after initialization."""
        self._validate_opacity()
        self._validate_content()
        self._validate_paths()

    def _validate_opacity(self):
        """Ensures opacity is within 0.0 to 1.0 range."""
        if not (0.0 <= self.opacity <= 1.0):
            raise InvalidInputError(f"Opacity must be between 0.0 and 1.0, got {self.opacity}")

    def _validate_content(self):
        """Ensures the correct content is provided for the selected type."""
        if self.watermark_type == WatermarkType.TEXT:
            if not self.text:
                raise InvalidInputError("Watermark type is TEXT, but 'text' content is missing.")
            if self.image_path:
                # Warning could be logged here, but we strictly ignore image path in text mode
                pass
                
        elif self.watermark_type == WatermarkType.IMAGE:
            if not self.image_path:
                raise InvalidInputError("Watermark type is IMAGE, but 'image_path' is missing.")

    def _validate_paths(self):
        """Checks if referenced files exist."""
        if self.image_path:
            path_obj = Path(self.image_path)
            if not path_obj.exists() or not path_obj.is_file():
                raise InvalidInputError(f"Image file not found at: {self.image_path}")
            self.image_path = path_obj  # Standardize to Path object
