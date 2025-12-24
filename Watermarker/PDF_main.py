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

import argparse
from PDFProcessor import PDFProcessor
from WatermarkConfig import WatermarkConfig
from WatermarkRenderer import WatermarkRenderer
# ==========================================
# CLI & Execution
# ==========================================

def run_watermark_service(
    input_pdf: str,
    output_pdf: str,
    watermark_text: Optional[str] = None,
    watermark_image: Optional[str] = None,
    position: str = "center",
    opacity: float = 0.5,
    rotation: float = 0.0,
    pages: Optional[str] = None,
    password: Optional[str] = None
):
    """
    High-level entry point to initialize and run the watermarking process.
    """
    # 1. Determine Type
    if watermark_image:
        w_type = WatermarkType.IMAGE
    else:
        w_type = WatermarkType.TEXT

    # 2. Map Position String to Enum
    pos_map = {p.value: p for p in WatermarkPosition}
    selected_pos = pos_map.get(position.lower(), WatermarkPosition.CENTER)

    # 3. Create Config
    config = WatermarkConfig(
        watermark_type=w_type,
        text=watermark_text,
        image_path=watermark_image,
        position=selected_pos,
        opacity=opacity,
        rotation=rotation,
        image_rel_page=0.3 if watermark_image else 0.0 # Default 30% width for images
    )

    # 4. Process
    processor = PDFProcessor(input_pdf, output_pdf, config, password=password)
    processor.process(page_selection=pages)
    processor.save_pdf()

def main():
    parser = argparse.ArgumentParser(
        description="Production PDF Watermark Adder (Text/Image)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    # Files
    parser.add_argument("-i", "--input", required=True, help="Path to source PDF")
    parser.add_argument("-o", "--output", required=True, help="Path to save watermarked PDF")
    
    # Watermark Content
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("-t", "--text", help="Text to use as watermark")
    group.add_argument("-img", "--image", help="Path to image (PNG/JPG) to use as watermark")
    
    # Appearance
    parser.add_argument("--pos", default="center", 
                        choices=["center", "top_left", "top_right", "bottom_left", "bottom_right"],
                        help="Position on the page")
    parser.add_argument("--opacity", type=float, default=0.3, help="Opacity (0.0 to 1.0)")
    parser.add_argument("--rotate", type=float, default=45.0, help="Rotation in degrees")
    parser.add_argument("--pages", help="Page range (e.g., '0, 2-4, 10')")
    parser.add_argument("--password", help="Password for encrypted PDFs")

    args = parser.parse_args()

    try:
        run_watermark_service(
            input_pdf=args.input,
            output_pdf=args.output,
            watermark_text=args.text,
            watermark_image=args.image,
            position=args.pos,
            opacity=args.opacity,
            rotation=args.rotate,
            pages=args.pages,
            password=args.password
        )
    except WatermarkError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    # Example for Library Usage:
    """
    config = WatermarkConfig(
        watermark_type=WatermarkType.TEXT,
        text="CONFIDENTIAL",
        opacity=0.2,
        rotation=45,
        position=WatermarkPosition.CENTER
    )
    proc = PDFProcessor("source.pdf", "output.pdf", config)
    proc.process(page_selection="0-2") # Only first 3 pages
    proc.save_pdf()
    """
    
    # Run CLI
    main()
