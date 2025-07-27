# AI Agent Development Guidelines

## Project Overview

The goal of this project is to build a template that allows AI agents to understand and autonomously operate on the content of Excel files.

## How to Use This Repository

### 1. File Preparation

Place the Excel file (.xlsx) you want to operate on in the `input` directory.

### 2. Content Conversion

Run the `xlsx_to_csv.py` script. This will convert each sheet of the Excel file into a separate CSV file and save them in the `csv_output` directory.

### 3. AI Understanding and Code Generation

As an AI agent, you will read the CSV files in the `csv_output` directory to understand the structure and content of the Excel file. Based on the user's instructions, you will then write the specific Excel operations you want to perform as Python code in the `custom_operation.py` file.

### 4. Execution

When `main.py` is executed, the code you generated in `custom_operation.py` will be run, and the original Excel file will be automatically edited.

## Your Role

Your primary responsibility is to **edit `custom_operation.py`** based on the user's request.

- **Analyze the CSV files** in the `csv_output` directory to understand the structure and content of the Excel file.
- **Identify the target cells** for the operation and determine the appropriate values or formulas to insert.
- **Write Python code** in `custom_operation.py` to perform the requested Excel operations.
  - Use the `openpyxl` library to read, modify, and write `.xlsx` files.
  - The relationship between the CSV data and the Excel sheet is as follows: the first row and column in the CSV correspond to the Excel sheet structure. Use this to calculate the correct cell references.

## Example Task: Calculating a Total

If the user asks you to "calculate the total sales", you should:

1.  **Analyze `csv_output/product_list.csv`** to identify the "Sales" column and the empty cell where the total should be placed.
2.  **Determine the range of cells** to be summed (e.g., `D2:D3` if "Sales" is in column D).
3.  **Write the following code** in `custom_operation.py`:

```python
from openpyxl import load_workbook

def custom_operation():
    wb = load_workbook('input/example.xlsx')
    sheet = wb['product_list']
    sheet['D4'] = '=SUM(D2:D3)'
    wb.save('input/example.xlsx')
```

This will insert the sum formula into the correct cell of the original Excel file.
