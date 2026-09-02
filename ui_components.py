import os
import sys

from PySide6.QtCore import Qt, QThread, Signal, QUrl
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QTextEdit, QFileDialog, QFrame,
    QGroupBox, QLineEdit, QMessageBox
)
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QDesktopServices

from converter_engine import convert_materials_to_pdf

class DropZoneWidget(QFrame):
    """Custom Drag and Drop zone."""
    files_dropped = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        self.icon_label = QLabel("[ DROP SOURCE FOLDER ]", self)
        self.icon_label.setObjectName("DropZoneIconText")
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        self.text_label = QLabel("Drag & Drop your Materials Folder here", self)
        self.text_label.setObjectName("DropZoneText")
        self.text_label.setAlignment(Qt.AlignCenter)
        
        self.subtext_label = QLabel("Supports PowerPoint (.pptx) and PDF (.pdf) lecture files", self)
        self.subtext_label.setObjectName("DropZoneSubtext")
        self.subtext_label.setAlignment(Qt.AlignCenter)
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.subtext_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setProperty("hover", "true")
            self.style().unpolish(self)
            self.style().polish(self)

    def dragLeaveEvent(self, event):
        self.setProperty("hover", "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event: QDropEvent):
        self.setProperty("hover", "false")
        self.style().unpolish(self)
        self.style().polish(self)

        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if os.path.exists(file_path):
                self.files_dropped.emit(file_path)

class WorkerThread(QThread):
    """Background worker thread for extraction & ReportLab PDF building."""
    progress_updated = Signal(int)
    log_emitted = Signal(str)
    conversion_finished = Signal(str)
    conversion_failed = Signal(str)

    def __init__(self, input_path, output_pdf_path):
        super().__init__()
        self.input_path = input_path
        self.output_pdf_path = output_pdf_path

    def run(self):
        try:
            out_pdf = convert_materials_to_pdf(
                self.input_path,
                self.output_pdf_path,
                progress_callback=self.progress_updated.emit,
                log_callback=self.log_emitted.emit
            )
            self.conversion_finished.emit(out_pdf)
        except Exception as e:
            self.conversion_failed.emit(str(e))

class ReviewerConverterWindow(QMainWindow):
    """Main Application Window."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Lesson-to-Reviewer Converter")
        self.resize(750, 680)
        self.setMinimumSize(650, 580)
        
        self.selected_input_path = ""
        self.last_output_pdf = ""

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(22, 22, 22, 22)
        main_layout.setSpacing(12)

        # Header Title
        title_label = QLabel("Lesson-to-Reviewer Converter", self)
        title_label.setObjectName("HeaderTitle")
        subtitle_label = QLabel("Convert lecture presentations and PDFs into a formatted 2-column Exam Reviewer PDF", self)
        subtitle_label.setObjectName("HeaderSubtitle")
        
        main_layout.addWidget(title_label)
        main_layout.addWidget(subtitle_label)

        # Drag and Drop Box
        self.drop_zone = DropZoneWidget(self)
        self.drop_zone.files_dropped.connect(self.on_input_selected)
        main_layout.addWidget(self.drop_zone)

        # Input Path Selection Bar
        input_box = QGroupBox("Selected Input Materials", self)
        input_layout = QHBoxLayout(input_box)
        
        self.input_path_edit = QLineEdit(self)
        self.input_path_edit.setPlaceholderText("No materials folder selected...")
        self.input_path_edit.setReadOnly(True)
        
        self.browse_folder_btn = QPushButton("Browse Folder", self)
        self.browse_folder_btn.clicked.connect(self.browse_folder)
        
        input_layout.addWidget(self.input_path_edit)
        input_layout.addWidget(self.browse_folder_btn)
        main_layout.addWidget(input_box)

        # Output Path Selection Bar
        output_box = QGroupBox("Output Location", self)
        output_layout = QHBoxLayout(output_box)
        
        self.output_path_edit = QLineEdit(self)
        default_out = os.path.join(os.path.expanduser("~"), "Downloads", "Generated_Midterm_Reviewer.pdf")
        self.output_path_edit.setText(default_out)
        
        self.browse_output_btn = QPushButton("Change Output...", self)
        self.browse_output_btn.clicked.connect(self.browse_output)
        
        output_layout.addWidget(self.output_path_edit)
        output_layout.addWidget(self.browse_output_btn)
        main_layout.addWidget(output_box)

        # Action Convert Button
        self.convert_btn = QPushButton("Generate Reviewer PDF", self)
        self.convert_btn.setObjectName("PrimaryButton")
        self.convert_btn.clicked.connect(self.start_conversion)
        main_layout.addWidget(self.convert_btn)

        # Progress Bar & Terminal Log
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.log_terminal = QTextEdit(self)
        self.log_terminal.setReadOnly(True)
        self.log_terminal.setMaximumHeight(130)
        self.log_terminal.append("System ready. Drag & drop a materials folder to begin.")
        main_layout.addWidget(self.log_terminal)

        # Completion Action Bar
        action_layout = QHBoxLayout()
        self.open_pdf_btn = QPushButton("Preview Generated Reviewer", self)
        self.open_pdf_btn.setObjectName("AccentButton")
        self.open_pdf_btn.setEnabled(False)
        self.open_pdf_btn.clicked.connect(self.open_generated_pdf)

        self.open_folder_btn = QPushButton("Open Output Folder", self)
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self.open_output_folder)

        action_layout.addWidget(self.open_pdf_btn)
        action_layout.addWidget(self.open_folder_btn)
        main_layout.addLayout(action_layout)

    def browse_folder(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Materials Directory")
        if dir_path:
            self.on_input_selected(dir_path)

    def browse_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Reviewer PDF", self.output_path_edit.text(), "PDF Files (*.pdf)")
        if file_path:
            self.output_path_edit.setText(file_path)

    def on_input_selected(self, path):
        self.selected_input_path = path
        self.input_path_edit.setText(path)
        self.drop_zone.text_label.setText(f"Selected: {os.path.basename(path)}")
        self.log_terminal.append(f"Selected source: {path}")

    def start_conversion(self):
        input_path = self.input_path_edit.text().strip()
        output_pdf = self.output_path_edit.text().strip()

        if not input_path or not os.path.exists(input_path):
            QMessageBox.warning(self, "Missing Input", "Please select or drop a valid materials folder first.")
            return

        if not output_pdf:
            QMessageBox.warning(self, "Missing Output Path", "Please specify a valid destination PDF path.")
            return

        self.convert_btn.setEnabled(False)
        self.browse_folder_btn.setEnabled(False)
        self.browse_output_btn.setEnabled(False)
        self.open_pdf_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_terminal.clear()
        self.log_terminal.append("Starting conversion process...")

        self.worker = WorkerThread(input_path, output_pdf)
        self.worker.progress_updated.connect(self.progress_bar.setValue)
        self.worker.log_emitted.connect(self.log_terminal.append)
        self.worker.conversion_finished.connect(self.on_conversion_success)
        self.worker.conversion_failed.connect(self.on_conversion_error)
        self.worker.start()

    def on_conversion_success(self, output_pdf_path):
        self.last_output_pdf = output_pdf_path
        self.progress_bar.setValue(100)
        self.log_terminal.append("\n[SUCCESS] CONVERSION COMPLETE!")
        self.log_terminal.append(f"Reviewer saved to: {output_pdf_path}")
        
        self.convert_btn.setEnabled(True)
        self.browse_folder_btn.setEnabled(True)
        self.browse_output_btn.setEnabled(True)
        self.open_pdf_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)

        QMessageBox.information(self, "Success", "Reviewer PDF generated successfully!")

    def on_conversion_error(self, error_msg):
        self.progress_bar.setValue(0)
        self.log_terminal.append(f"\n[ERROR]: {error_msg}")
        
        self.convert_btn.setEnabled(True)
        self.browse_folder_btn.setEnabled(True)
        self.browse_output_btn.setEnabled(True)

        QMessageBox.critical(self, "Conversion Error", f"Failed to generate reviewer:\n{error_msg}")

    def open_generated_pdf(self):
        if self.last_output_pdf and os.path.exists(self.last_output_pdf):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.last_output_pdf))

    def open_output_folder(self):
        if self.last_output_pdf:
            folder = os.path.dirname(self.last_output_pdf)
            if os.path.exists(folder):
                QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
