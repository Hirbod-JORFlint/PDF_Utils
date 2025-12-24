import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFileDialog, QLineEdit, QComboBox, 
                             QDoubleSpinBox, QFrame, QRadioButton, QButtonGroup, QGroupBox)
from PyQt6.QtGui import QPixmap, QImage, QPainter, QColor, QFont, QFontMetrics, QTransform
from PyQt6.QtCore import Qt

# Import your existing logic
from PDF_main import run_watermark_service

class WatermarkGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pro PDF Watermarker")
        self.setMinimumSize(1100, 750)
        
        # State variables
        self.input_path = ""
        self.watermark_image_path = None
        
        self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout()
        
        # ===========================
        # Left Panel: Controls
        # ===========================
        controls_widget = QWidget()
        controls_layout = QVBoxLayout()
        controls_widget.setFixedWidth(400)

        # --- 1. File Selection Section ---
        file_group = QGroupBox("Source File")
        file_layout = QVBoxLayout()
        
        self.btn_browse = QPushButton("Select Source PDF")
        self.btn_browse.clicked.connect(self.open_file)
        self.lbl_file = QLabel("No file selected")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setStyleSheet("color: #666; font-style: italic;")
        
        self.txt_password = QLineEdit()
        self.txt_password.setPlaceholderText("PDF Password (if encrypted)")
        self.txt_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_password.textChanged.connect(self.update_preview)

        file_layout.addWidget(self.btn_browse)
        file_layout.addWidget(self.lbl_file)
        file_layout.addWidget(self.txt_password)
        file_group.setLayout(file_layout)

        # --- 2. Watermark Content Section ---
        content_group = QGroupBox("Watermark Content")
        content_layout = QVBoxLayout()

        type_layout = QHBoxLayout()
        self.radio_text = QRadioButton("Text")
        self.radio_image = QRadioButton("Image")
        self.radio_text.setChecked(True)
        
        self.radio_text.toggled.connect(self.toggle_content_mode)
        self.radio_image.toggled.connect(self.toggle_content_mode)
        
        # Connect radios to update preview immediately when switched
        self.radio_text.toggled.connect(self.update_preview)
        self.radio_image.toggled.connect(self.update_preview)
        
        type_layout.addWidget(self.radio_text)
        type_layout.addWidget(self.radio_image)

        self.txt_watermark = QLineEdit()
        self.txt_watermark.setPlaceholderText("Enter Watermark Text...")
        self.txt_watermark.textChanged.connect(self.update_preview)

        self.btn_img_browse = QPushButton("Select Watermark Image")
        self.btn_img_browse.clicked.connect(self.select_watermark_image)
        self.btn_img_browse.setVisible(False)
        
        self.lbl_img_path = QLabel("No image selected")
        self.lbl_img_path.setStyleSheet("color: #666; font-size: 10px;")
        self.lbl_img_path.setVisible(False)

        content_layout.addLayout(type_layout)
        content_layout.addWidget(self.txt_watermark)
        content_layout.addWidget(self.btn_img_browse)
        content_layout.addWidget(self.lbl_img_path)
        content_group.setLayout(content_layout)

        # --- 3. Appearance Settings ---
        settings_group = QGroupBox("Appearance & Scope")
        settings_layout = QVBoxLayout()

        settings_layout.addWidget(QLabel("Position:"))
        self.combo_pos = QComboBox()
        self.combo_pos.addItems(["center", "top_left", "top_right", "bottom_left", "bottom_right"])
        self.combo_pos.currentTextChanged.connect(self.update_preview)

        settings_layout.addWidget(QLabel("Opacity (0.1 - 1.0):"))
        self.spin_opacity = QDoubleSpinBox()
        self.spin_opacity.setRange(0.1, 1.0)
        self.spin_opacity.setValue(0.5)
        self.spin_opacity.setSingleStep(0.1)
        self.spin_opacity.valueChanged.connect(self.update_preview)
        
        settings_layout.addWidget(QLabel("Rotation (Degrees):"))
        self.spin_rotate = QDoubleSpinBox()
        self.spin_rotate.setRange(0, 360)
        self.spin_rotate.setValue(45.0)
        self.spin_rotate.valueChanged.connect(self.update_preview)

        settings_layout.addWidget(QLabel("Page Range (e.g., '0, 2-4'):"))
        self.txt_pages = QLineEdit()
        self.txt_pages.setPlaceholderText("Leave empty for all pages")
        self.txt_pages.textChanged.connect(self.update_preview)

        settings_layout.addWidget(self.combo_pos)
        settings_layout.addWidget(self.spin_opacity)
        settings_layout.addWidget(self.spin_rotate)
        settings_layout.addWidget(self.txt_pages)
        settings_group.setLayout(settings_layout)

        # --- Process Button ---
        self.btn_process = QPushButton("Apply Watermark & Save")
        self.btn_process.setStyleSheet("""
            QPushButton {
                background-color: #0078D7; 
                color: white; 
                font-weight: bold; 
                height: 45px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #0063B1; }
        """)
        self.btn_process.clicked.connect(self.save_pdf)

        controls_layout.addWidget(file_group)
        controls_layout.addWidget(content_group)
        controls_layout.addWidget(settings_group)
        controls_layout.addStretch()
        controls_layout.addWidget(self.btn_process)
        controls_widget.setLayout(controls_layout)

        # ===========================
        # Right Panel: Preview
        # ===========================
        self.preview_area = QLabel("Select a PDF to see preview")
        self.preview_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_area.setStyleSheet("border: 2px dashed #ccc; background-color: #f0f0f0; color: #888;")

        main_layout.addWidget(controls_widget)
        main_layout.addWidget(self.preview_area, 1)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    # ===========================
    # Logic Methods
    # ===========================

    def toggle_content_mode(self):
        is_text = self.radio_text.isChecked()
        self.txt_watermark.setVisible(is_text)
        self.btn_img_browse.setVisible(not is_text)
        self.lbl_img_path.setVisible(not is_text)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.input_path = file_path
            self.lbl_file.setText(f"Selected: {file_path.split('/')[-1]}")
            self.update_preview()

    def select_watermark_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.png *.jpg *.jpeg)")
        if file_path:
            self.watermark_image_path = file_path
            self.lbl_img_path.setText(f"Selected: {file_path.split('/')[-1]}")
            self.update_preview()

    def update_preview(self):
        """Generates a preview with the simulated watermark overlay."""
        if not self.input_path:
            return

        try:
            doc = fitz.open(self.input_path)
            
            # --- 1. Password Handling ---
            if doc.needs_pass:
                pwd = self.txt_password.text()
                if not doc.authenticate(pwd):
                    self.preview_area.setText("🔒 PDF is Locked.\nEnter Password to Preview.")
                    doc.close()
                    return

            # --- 2. Page Selection Logic ---
            page_index = 0
            page_text = self.txt_pages.text().strip()
            if page_text:
                try:
                    first_group = page_text.split(',')[0].strip()
                    first_num_str = first_group.split('-')[0].strip()
                    if first_num_str.isdigit():
                        val = int(first_num_str)
                        if val > 0: page_index = val - 1
                except ValueError:
                    pass

            total_pages = doc.page_count
            if page_index >= total_pages:
                self.preview_area.setText(f"Page {page_index + 1} out of range\n(Doc has {total_pages} pages)")
                doc.close()
                return

            # --- 3. Render PDF Page to Pixmap ---
            page = doc.load_page(page_index)
            # Use 1.0 scale for clearer text rendering before scaling down for display
            pix = page.get_pixmap(matrix=fitz.Matrix(1.0, 1.0))
            
            fmt = QImage.Format.Format_RGB888
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
            canvas_pixmap = QPixmap.fromImage(qimg)
            
            doc.close()

            # --- 4. Draw Watermark Overlay (QPainter) ---
            painter = QPainter(canvas_pixmap)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # Get UI settings
            opacity = self.spin_opacity.value()
            rotation = self.spin_rotate.value()
            position_str = self.combo_pos.currentText()
            
            # Calculate Center Coordinates based on Position
            # We determine where the CENTER of the watermark should be placed
            w, h = canvas_pixmap.width(), canvas_pixmap.height()
            margin_x = w * 0.15
            margin_y = h * 0.15
            
            if position_str == "center":
                cx, cy = w / 2, h / 2
            elif position_str == "top_left":
                cx, cy = margin_x, margin_y
            elif position_str == "top_right":
                cx, cy = w - margin_x, margin_y
            elif position_str == "bottom_left":
                cx, cy = margin_x, h - margin_y
            elif position_str == "bottom_right":
                cx, cy = w - margin_x, h - margin_y
            else:
                cx, cy = w / 2, h / 2

            # Translate painter to the target point so we can rotate around it
            painter.translate(cx, cy)
            painter.rotate(-rotation) # Negative because QPainter rotates clockwise, Math/PDF usually counter-clockwise

            # Draw Content (Text or Image)
            if self.radio_text.isChecked():
                wm_text = self.txt_watermark.text()
                if wm_text:
                    # Dynamic font size (approx 1/20th of page width)
                    font_size = max(12, int(w / 20)) 
                    font = QFont("Helvetica", font_size, QFont.Weight.Bold)
                    painter.setFont(font)
                    
                    # Set Opacity via Alpha Channel (0-255)
                    color = QColor(80, 80, 80, int(opacity * 255)) # Dark Grey
                    painter.setPen(color)
                    
                    # Center text at (0,0) (which is now cx, cy)
                    metrics = QFontMetrics(font)
                    rect = metrics.boundingRect(wm_text)
                    text_w, text_h = rect.width(), rect.height()
                    painter.drawText(-text_w // 2, text_h // 4, wm_text)

            else:
                # Image Mode
                if self.watermark_image_path:
                    wm_pix = QPixmap(self.watermark_image_path)
                    if not wm_pix.isNull():
                        # Scale image to ~30% of page width
                        target_w = int(w * 0.3)
                        wm_pix = wm_pix.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
                        
                        painter.setOpacity(opacity)
                        img_w, img_h = wm_pix.width(), wm_pix.height()
                        painter.drawPixmap(-img_w // 2, -img_h // 2, wm_pix)

            painter.end()

            # --- 5. Display Result ---
            self.lbl_file.setText(f"Previewing Page {page_index + 1} of {total_pages}")
            self.preview_area.setPixmap(canvas_pixmap.scaled(
                self.preview_area.width(), 
                self.preview_area.height(), 
                Qt.AspectRatioMode.KeepAspectRatio, 
                Qt.TransformationMode.SmoothTransformation
            ))
            
        except Exception as e:
            self.preview_area.setText(f"Error loading preview: {e}")
            print(e) # For debugging in console

    def save_pdf(self):
        if not self.input_path:
            self.lbl_file.setText("Error: No PDF selected!")
            return
            
        output_path, _ = QFileDialog.getSaveFileName(self, "Save Watermarked PDF", "", "PDF Files (*.pdf)")
        if not output_path:
            return

        wm_text = None
        wm_image = None
        
        if self.radio_text.isChecked():
            wm_text = self.txt_watermark.text()
            if not wm_text:
                self.lbl_file.setText("Error: Enter watermark text!")
                return
        else:
            wm_image = self.watermark_image_path
            if not wm_image:
                self.lbl_file.setText("Error: Select an image!")
                return

        page_selection = self.txt_pages.text().strip() or None
        pdf_password = self.txt_password.text().strip() or None

        try:
            run_watermark_service(
                input_pdf=self.input_path,
                output_pdf=output_path,
                watermark_text=wm_text,
                watermark_image=wm_image,
                position=self.combo_pos.currentText(),
                opacity=self.spin_opacity.value(),
                rotation=self.spin_rotate.value(),
                pages=page_selection,
                password=pdf_password
            )
            self.lbl_file.setText("✅ Success! PDF Saved.")
        except Exception as e:
            self.lbl_file.setText(f"❌ Error: {e}")
            print(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WatermarkGUI()
    window.show()
    sys.exit(app.exec())
