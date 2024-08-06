import sys
from PyQt5 import uic
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from detect import *
from time import sleep

Ui_MainWindow, QtBaseClass = uic.loadUiType('converter.ui')

class Worker(QThread):
    finished = pyqtSignal()
    alert_signal = pyqtSignal(str)

    def __init__(self, file_path, output_file_name, column_list, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.output_file_name = output_file_name
        self.column_list = column_list

    def run(self):
        try:
            table_df = extract_table_from_excel(self.file_path)
            print(table_df)

            complete_column(table_df, self.column_list[0], clean_date_format, 1)
            complete_column(table_df, self.column_list[1], clean_time_format, 2)
            complete_column(table_df, self.column_list[2], clean_symbol_format, 3)
            complete_column(table_df, self.column_list[3], clean_quantity_format, 4)
            complete_column(table_df, self.column_list[4], clean_price_format, 5)
            complete_column(table_df, self.column_list[5], clean_commission_format, 6)
            complete_column(table_df, self.column_list[6], clean_action_format, 7)
            
            output_file_path = self.output_file_name + ".xlsx"
            wb.save(output_file_path)

            generate_final_result(output_file_path)

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
        
        file_path = self.ui.files.item(0).text()
        
        # Check if the file is an Excel or CSV file
        if not (file_path.endswith('.xlsx') or file_path.endswith('.xls') or file_path.endswith('.csv')):
            QMessageBox.warning(self, "Warning", "Unsupported file type! Please import an Excel or CSV file.")
            self.ui.files.clear()
            return
        
        if self.ui.output_name.text() == "":
            QMessageBox.warning(self, "Warning", "Please enter the output file name !")
            return

        output_file_name = self.ui.output_name.text()
        
        self.ui.convert_btn.setEnabled(False)
        self.ui.import_btn.setEnabled(False)

        sleep(1)

        column_list = [self.ui.date_column.text(), self.ui.time_column.text(), self.ui.symbol_column.text(), self.ui.quantity_column.text(), self.ui.price_column.text(), self.ui.commission_column.text(), self.ui.action_column.text()]
        print(column_list)

        self.worker = Worker(file_path, output_file_name, column_list)
        self.worker.finished.connect(self.on_finished)
        self.worker.alert_signal.connect(self.show_alert)
        self.worker.start()

    def on_finished(self):
        # Clear all items in the QListWidget
        self.ui.files.clear()
        self.ui.output_name.clear()
        self.ui.date_column.clear()
        self.ui.time_column.clear()
        self.ui.symbol_column.clear()
        self.ui.quantity_column.clear()
        self.ui.price_column.clear()
        self.ui.commission_column.clear()
        self.ui.action_column.clear()

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