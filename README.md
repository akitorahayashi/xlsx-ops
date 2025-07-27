# Excel to CSV Converter

This is a template project designed for an AI agent to read and autonomously manipulate Excel data. The ultimate goal is for the AI agent to understand the contents of an .xlsx file and automatically generate Python code to edit the file based on user instructions.

## Directory Structure

```text
/
|-- csv_output/
|   |-- xlsx_to_csv.py
|   |-- sheet1.csv
|   `-- product_list.csv
|-- input/
|   `-- example.xlsx
|-- custom_operation.py
|-- main.py
...
```

- **`csv_output/xlsx_to_csv.py`**: The Python script that converts Excel files to CSV.
- **`csv_output/`**: The directory where the output CSV files are saved.
- **`input/`**: The directory where you should place your `.xlsx` files.
- **`custom_operation.py`**: A script where the AI agent writes the Python code to manipulate the Excel file.
- **`main.py`**: The main script to run the entire process.

## Setup

### Prerequisites

- Python 3.9 or higher
- `pip` for package installation

### Installation

Install the required Python packages:
```bash
pip install pandas openpyxl
```

## How to Use

Place your Excel (.xlsx) file in the input directory.

Run the `csv_output/xlsx_to_csv.py` script to convert each sheet into a separate CSV file in the same directory.

Review the generated CSV files to understand the structure of your Excel data.

Provide instructions to the AI agent on how you want to modify the .xlsx file. For example: "In the sheet named 'Orders', calculate the sum of the 'Order Amount' column and place the total in the 'Total' section."

The AI agent will analyze the CSV files to understand the .xlsx structure and then write the necessary Python code into custom_operation.py to perform the requested modifications. Note: The AI agent must only modify custom_operation.py and must not alter main.py.

Run main.py to execute the code in custom_operation.py.

Find the modified .xlsx file in the output directory.

## CSV Format

The script uses the `pandas` library to convert each sheet into CSV format.

- **Formulas**: Formulas in the Excel file will be output as their calculated values.
- **Cell Positioning**: The first row and column in the CSV correspond to the Excel sheet structure. The AI agent uses this correspondence to interpret the cell positions.

Example output:
```csv
Product,Price,Sales
Apple,150,30
Orange,100,50
Total,,80
```

## Expected Benefits

- **Data Structure Comprehension**: The AI agent can intuitively understand the data structure of each sheet.
- **Rational Decision Making**: By understanding the context of the data, the AI agent can perform reasonable operations based on abstract instructions.
