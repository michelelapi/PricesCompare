# PricesCompare

A Python GUI application to compare prices of the same items across multiple Excel files.

## Features
- Select multiple Excel files to compare
- For each file, choose which columns represent the item name and the price
- Compares all items and finds the lowest price for each
- Displays the results in the application
- Save the comparison results as a CSV file

## Requirements
- Python 3.7+
- pandas
- openpyxl
- Tkinter (comes with standard Python)

## Installation
1. Clone this repository or download the source code.
2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```
3. (Optional) Use the provided batch script to run the app and handle dependencies:
   ```
   run_price_compare.bat
   ```

## Usage
1. Run the application:
   ```
   python price_compare_gui.py
   ```
2. Click "Select Excel Files" and choose the files you want to compare.
3. For each file, select the item and price columns.
4. The app will display the items with the lowest price found across all files.
5. Click "Save Results as CSV" to export the results.

## Author
- [michelelapi](https://github.com/michelelapi)

---
Feel free to contribute or open issues for suggestions and bug reports! 