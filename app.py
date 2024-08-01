import sys
from PyQt6 import uic
from PyQt6.QtCore import QThread, pyqtSignal
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
import threading
from detect import *
from time import sleep

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
        if file_names:
            self.ui.files.clear()
            self.ui.files.addItem(file_names[0])

    def run_script(self):
        if self.ui.files.count() == 0:
            QMessageBox.warning(self, "Warning", "Import *.xlsx or *.csv file !")
            return
        
        if self.ui.output_name.text() == "":
            QMessageBox.warning(self, "Warning", "Enter output file name !")
            return

        self.file_path = self.ui.files.item(0).text()
        self.output_file_name = self.ui.output_name.text()

        self.ui.convert_btn.setEnabled(False)
        self.ui.import_btn.setEnabled(False)

        sleep(3)

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

        self.ui.files.clear()
        self.ui.output_name.clear()

        self.ui.convert_btn.setEnabled(True)
        self.ui.import_btn.setEnabled(True)
        
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