import sys
from PyQt6 import uic
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog
from openpyxl import Workbook
import threading

wb = Workbook()
sheet = wb.active

Ui_MainWindow, QtBaseClass = uic.loadUiType('converter.ui')

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        # Connect button click event to a function
        self.ui.import_btn.clicked.connect(self.upload_files)
        self.ui.convert_btn.clicked.connect(self.run_script_in_thread)

        self.ui.convert_btn.setEnabled(True)
        self.ui.import_btn.setEnabled(True)

    def upload_files(self):
        file_names, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", "All Files (*)")
        self.ui.files.addItems(file_names)

    def run_script(self):
        self.ui.convert_btn.setEnabled(False)
        self.ui.import_btn.setEnabled(False)

        self.file_path = self.ui.files.item(0).text()
        print(self.file_path)

    def run_script_in_thread(self):
        thread = threading.Thread(target = self.run_script)
        thread.start()

    def open(self):
        self.open()

    def reject(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())