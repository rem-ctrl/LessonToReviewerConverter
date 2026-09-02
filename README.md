# Lesson-to-Reviewer Converter

A modern **PySide6 (Qt 6)** Desktop Application & Standalone Executable that converts lecture presentation folders (.pptx and .pdf) into formatted **2-column Exam Reviewer PDFs**.

![Application UI Screenshot](app_preview.png)

## Features
- **Drag & Drop Box**: Drop any folder containing PowerPoint (.pptx) or PDF (.pdf) lecture materials directly into the app window.
- **Multi-Threaded Conversion Engine**: Runs file extraction and ReportLab PDF building on a background worker thread (QThread).
- **Automated Highlighting & Comparative Matrices**: Wraps key terms/names/dates in yellow highlights, exam traps in red warnings, and builds side-by-side comparative matrices.
- **Standalone Windows Executable (.exe)**: Ready to run portable binary without requiring Python.

## Project Structure
`
LessonToReviewerConverter/
├── app.py                 # Main GUI Application Entry Point
├── ui_components.py       # PySide6 Widgets & Drag-and-Drop Dropzone
├── converter_engine.py    # Text Extraction & ReportLab PDF Generator
├── styles.qss             # QSS Dark Theme Stylesheet
└── dist/
    └── LessonToReviewerConverter.exe  # Portable Executable
`

## Quick Start

### Option 1: Run Standalone Executable
Double-click dist/LessonToReviewerConverter.exe.

### Option 2: Run via Python
1. Install dependencies:
   `ash
   pip install PySide6 python-pptx pypdf reportlab
   `
2. Run the application:
   `ash
   python app.py
   `
