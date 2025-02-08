from PyQt6.QtWidgets import QApplication
from certificate_app import CertificateApp

if __name__ == "__main__":
    app = QApplication([])
    window = CertificateApp()
    window.show()
    app.exec()
