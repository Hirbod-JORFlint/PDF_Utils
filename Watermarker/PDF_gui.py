import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QFileDialog, QLineEdit, QComboBox, 
                             QDoubleSpinBox, QFrame, QRadioButton, QButtonGroup, QGroupBox)
from PyQt6.QtGui import QPixmap, QImage
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
        controls_widget.setFixedWidth(400) # Slightly wider for new options

        # --- 1. File Selection Section ---
        file_group = QGroupBox("Source File")
        file_layout = QVBoxLayout()
        
        self.btn_browse = QPushButton("Select Source PDF")
        self.btn_browse.clicked.connect(self.open_file)
        self.lbl_file = QLabel("No file selected")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setStyleSheet("color: #666; font-style: italic;")
        
        # Password Field (Hidden by default, useful for previewing locked files)
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

        # Type Selection (Radio Buttons)
        type_layout = QHBoxLayout()
        self.radio_text = QRadioButton("Text")
        self.radio_image = QRadioButton("Image")
        self.radio_text.setChecked(True)
        
        self.radio_text.toggled.connect(self.toggle_content_mode)
        self.radio_image.toggled.connect(self.toggle_content_mode)
        
        type_layout.addWidget(self.radio_text)
        type_layout.addWidget(self.radio_image)

        # Input Fields
        self.txt_watermark = QLineEdit()
        self.txt_watermark.setPlaceholderText("Enter Watermark Text...")
        self.txt_watermark.textChanged.connect(self.update_preview) # Refresh preview on text change

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

        # Position
        settings_layout.addWidget(QLabel("Position:"))
        self.combo_pos = QComboBox()
        self.combo_pos.addItems(["center", "top_left", "top_right", "bottom_left", "bottom_right"])

        # Opacity
        settings_layout.addWidget(QLabel("Opacity (0.1 - 1.0):"))
        self.spin_opacity = QDoubleSpinBox()
        self.spin_opacity.setRange(0.1, 1.0)
        self.spin_opacity.setValue(0.5)
        self.spin_opacity.setSingleStep(0.1)
        
        # Rotation
        settings_layout.addWidget(QLabel("Rotation (Degrees):"))
        self.spin_rotate = QDoubleSpinBox()
        self.spin_rotate.setRange(0, 360)
        self.spin_rotate.setValue(45.0)

        # Page Range
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

        # Add all to left panel
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
        main_layout.addWidget(self.preview_area, 1) # Stretch factor 1

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget) # Fixed: Added self.

    # ===========================
    # Logic Methods
    # ===========================

    def toggle_content_mode(self):
        """Switches UI between Text and Image input modes."""
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

    def update_preview(self):
            """Generates a preview of the specific page selected in the range."""
            if not self.input_path:
                return

            try:
                doc = fitz.open(self.input_path)
                
                # 1. Handle Password
                if doc.needs_pass:
                    pwd = self.txt_password.text()
                    if not doc.authenticate(pwd):
                        self.preview_area.setText("🔒 PDF is Locked.\nEnter Password to Preview.")
                        doc.close()
                        return

                # 2. Determine which page to show
                page_index = 0 # Default to first page
                page_text = self.txt_pages.text().strip()
                
                if page_text:
                    try:
                        # Logic: Get the first number from strings like "3-5, 10" or "2, 4"
                        # Step A: Split by comma to get first group ("3-5")
                        first_group = page_text.split(',')[0].strip()
                        # Step B: Split by dash to get start of range ("3")
                        first_num_str = first_group.split('-')[0].strip()
                        
                        if first_num_str.isdigit():
                            val = int(first_num_str)
                            # User input is 1-based, internal PDF is 0-based
                            if val > 0:
                                page_index = val - 1
                    except ValueError:
                        pass # Fallback to 0 if parsing fails

                # 3. Boundary Check (Prevent crashing if user types '100' for a 5-page doc)
                total_pages = doc.page_count
                if page_index >= total_pages:
                    self.preview_area.setText(f"Page {page_index + 1} out of range\n(Doc has {total_pages} pages)")
                    doc.close()
                    return

                # 4. Load & Render Page
                page = doc.load_page(page_index) 
                
                # Render at slightly higher quality
                pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8)) 
                
                fmt = QImage.Format.Format_RGB888
                qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
                pixmap = QPixmap.fromImage(qimg)
                
                # Display
                self.preview_area.setPixmap(pixmap.scaled(
                    self.preview_area.width(), 
                    self.preview_area.height(), 
                    Qt.AspectRatioMode.KeepAspectRatio, 
                    Qt.TransformationMode.SmoothTransformation
                ))
                
                # Update label to confirm which page is being shown
                self.lbl_file.setText(f"Previewing Page {page_index + 1} of {total_pages}")
                
                doc.close()
                
            except Exception as e:
                self.preview_area.setText(f"Error loading preview: {e}")
                
    def save_pdf(self):
        if not self.input_path:
            self.lbl_file.setText("Error: No PDF selected!")
            return
            
        output_path, _ = QFileDialog.getSaveFileName(self, "Save Watermarked PDF", "", "PDF Files (*.pdf)")
        if not output_path:
            return

        # Prepare arguments based on mode
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

        # Prepare Page Range (Convert empty string to None)
        page_selection = self.txt_pages.text().strip()
        if not page_selection:
            page_selection = None
            
        # Prepare Password (Convert empty string to None)
        pdf_password = self.txt_password.text().strip()
        if not pdf_password:
            pdf_password = None

        try:
            # Call the backend logic
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
