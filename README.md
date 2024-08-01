# Excel Converter Desktop App

## Overview
This desktop application allows users to convert and clean up data from Excel files (.xlsx, .xls) or CSV files into a standardized format. It is built using the PyQt6 framework for the GUI and utilizes pandas for data manipulation.

## Features
- **File Import:** Supports importing of Excel (.xlsx, .xls) and CSV files.
- **Data Cleaning:** Automatically cleans and standardizes various columns such as dates, times, symbols, quantities, prices, commissions, and actions.
- **Multithreading:** Performs conversion in a separate thread to keep the UI responsive.
- **Error Handling:** Provides informative alerts in case of issues during the conversion process.

## Installation

### Prerequisites
- Python 3.x
- pip (Python package installer)

### Required Libraries
You can install the required libraries using the following command:
```sh
pip install PyQt6 pandas openpyxl python-dateutil
```

## Usage

1. **Clone the repository:**
   ```sh
   git clone https://github.com/hi-tech-AI/excel-converter.git
   cd excel-converter
   ```

2. **Prepare the `.ui` file:**
   Ensure you have a Qt Designer-generated `converter.ui` file in the project directory.

3. **Run the application:**
   ```sh
   python app.py
   ```

4. **Using the Application:**
   - Click on the "Import" button to select an Excel or CSV file.
   - Enter the desired output file name.
   - Click on the "Convert" button to start the conversion process.
   - The application will notify you when the conversion is complete or if it encounters any errors.

## Project Structure

```
excel-converter-app/
├── detect.py         # Contains functions for extracting and cleaning data
├── app.py            # Main application file with GUI logic
├── converter.ui      # Qt Designer file defining the UI layout
├── README.md         # This readme file
├── requirements.txt  # List of required python packages
```

## Functions

### `detect.py`

- **`extract_table_from_excel(file_path)`**: Extracts tables from an Excel or CSV file.
- **`clean_date_format(date_string)`**: Cleans and formats date strings.
- **`clean_time_format(time_string)`**: Cleans and formats time strings.
- **`clean_symbol_format(symbol_string)`**: Cleans and formats symbol strings.
- **`clean_quantity_format(quantity_string)`**: Cleans and formats quantity strings.
- **`clean_price_format(price_string)`**: Cleans and formats price strings.
- **`clean_commission_format(commission_string)`**: Cleans and formats commission strings.
- **`clean_action_format(action_string)`**: Cleans and formats action strings.
- **`complete_column(table_df, search_key, function, column)`**: Completes a column by applying a specific cleaning function.

### `app.py`

- **Worker(QThread)**: A class that performs data extraction and cleaning in a separate thread.
- **MainWindow(QMainWindow)**: The main window class that handles user interactions and displays the GUI.

## Contributing

Feel free to submit issues or pull requests if you find bugs or think of improvements for the application.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Enjoy converting and cleaning your data effortlessly with this PyQt6-based desktop app!
