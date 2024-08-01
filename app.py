import sys
from PyQt6 import uic
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from detect import *
from time import sleep

Ui_MainWindow, QtBaseClass = uic.loadUiType('converter.ui')

class Worker(QThread):
    finished = pyqtSignal()
    alert_signal = pyqtSignal(str)

    def __init__(self, file_path, output_file_name, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.output_file_name = output_file_name

    def run(self):
        try:
            table_df = extract_table_from_excel(self.file_path)
            print(table_df)
            
            check_date_list = ['trade_date', 'TransactionDate', 'Date', 'TD']
            check_time_list = ['time', 'Activity Time']
            check_symbol_list = ['instrument.cns_equity_master.symbol', 'Symbol']
            check_quantity_list = ['quantity', 'Quantity']
            check_price_list = ['price', 'Price']
            check_commission_list = ['fee_commission', 'Commission', 'Fees & Comm']
            check_action_list = ['side_direction', 'TransactionType', 'Action', 'Transaction']
            complete_column(table_df, check_date_list, clean_date_format, 1)
            complete_column(table_df, check_time_list, clean_time_format, 2)
            complete_column(table_df, check_symbol_list, clean_symbol_format, 3)
            complete_column(table_df, check_quantity_list, clean_quantity_format, 4)
            complete_column(table_df, check_price_list, clean_price_format, 5)
            complete_column(table_df, check_commission_list, clean_commission_format, 6)
            complete_column(table_df, check_action_list, clean_action_format, 7)
            wb.save(f"{self.output_file_name}.xlsx")
        except Exception as e:
            self.alert_signal.emit(str(e))
        
        self.finished.emit()


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
        if file_names:
            # Clear the QListWidget before adding new items
            self.ui.files.clear()
            # Add only the first selected file to the QListWidget
            self.ui.files.addItem(file_names[0])

    def run_script_in_thread(self):
        # Check if QListWidget is empty
        if self.ui.files.count() == 0:
            QMessageBox.warning(self, "Warning", "Please import an Excel or CSV file.")
            return
        
        self.file_path = self.ui.files.item(0).text()
        
        # Check if the file is an Excel or CSV file
        if not (self.file_path.endswith('.xlsx') or self.file_path.endswith('.xls') or self.file_path.endswith('.csv')):
            QMessageBox.warning(self, "Warning", "Unsupported file type! Please import an Excel or CSV file.")
            self.ui.files.clear()
            return
        
        if self.ui.output_name.text() == "":
            QMessageBox.warning(self, "Warning", "Please enter the output file name !")
            return

        self.output_file_name = self.ui.output_name.text()
        
        self.ui.convert_btn.setEnabled(False)
        self.ui.import_btn.setEnabled(False)

        sleep(3)

        self.worker = Worker(self.file_path, self.output_file_name)
        self.worker.finished.connect(self.on_finished)
        self.worker.alert_signal.connect(self.show_alert)
        self.worker.start()

    def on_finished(self):
        # Clear all items in the QListWidget
        self.ui.files.clear()
        self.ui.output_name.clear()

        self.ui.convert_btn.setEnabled(True)
        self.ui.import_btn.setEnabled(True)

        QMessageBox.warning(self, "OK", "The file was successfully converted !")

    def show_alert(self, message):
        QMessageBox.warning(self, "Warning", message)

    def open(self):
        self.open()

    def reject(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())