# Excel to CSV Converter

This project provides a Python script to convert Excel files (.xlsx) into CSV files, making them readable and understandable for AI agents.

## Overview

The primary goal of this project is to enable AI agents to autonomously interact with Excel files. Since AI agents cannot directly read the binary content of `.xlsx` files, this script converts each sheet of an Excel file into a separate CSV file. This allows the agent to understand the structure and content of the Excel file, enabling it to perform data manipulation tasks.

## Directory Structure

```
/
|-- xlsx_to_csv.py
|-- csv_output/
|   |-- sheet1.csv
|   `-- product_list.csv
|-- input/
|   `-- example.xlsx
|-- custom_operation.py
|-- main.py
...
```

- **`xlsx_to_csv.py`**: The Python script that converts Excel files to CSV.
- **`csv_output/`**: The directory where the output CSV files are saved.
- **`input/`**: The directory where you should place your `.xlsx` files.
- **`custom_operation.py`**: A script where the AI agent will write code to manipulate the Excel file based on its understanding of the CSV files.
- **`main.py`**: The main script to run the entire process.

## How to Use

### Prerequisites

Install the required Python packages:
```bash
pip install pandas openpyxl
```

1. **Place your Excel file in the `input` directory.**
   - Make sure the file has a `.xlsx` extension.

2. **Run the conversion script.**
   ```bash
   python xlsx_to_csv.py
   ```

3. **Check the `csv_output` directory.**
   - The script will generate a CSV file for each sheet in the Excel file. The filename will be `{sheet_name}.csv`.

## CSV Format

The script uses the `pandas` library to convert each sheet into CSV format.

- **Formulas**: Formulas in the Excel file will be output as their calculated values.
- **Cell Positioning**: The first row and column in the CSV correspond to the Excel sheet structure. The AI agent uses this correspondence to interpret the cell positions.

Example output:
```
Product,Price,Sales
Apple,150,30
Orange,100,50
Total,,80
```

## Expected Benefits

- **Data Structure Comprehension**: The AI agent can intuitively understand the data structure of each sheet.
- **Rational Decision Making**: By understanding the context of the data, the AI agent can perform reasonable operations based on abstract instructions.
