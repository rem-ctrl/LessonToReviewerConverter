import os
import sys

from PySide6.QtWidgets import QApplication
from ui_components import ReviewerConverterWindow

def main():
    app = QApplication(sys.argv)
    
    # Load QSS Dark Theme stylesheet
    qss_path = os.path.join(os.path.dirname(__file__), "styles.qss")
    if os.path.exists(qss_path):
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    window = ReviewerConverterWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
