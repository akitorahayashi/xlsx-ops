# AI Agent Development Guidelines

## Project Overview

The goal of this project is to build a template that allows AI agents to understand and autonomously operate on the content of Excel files.

## Challenge to Solve

Currently, the system lacks the functionality for an AI agent to directly read the binary content of `.xlsx` files and understand their structure (e.g., sheet names, cell values, formulas). This prevents the agent from planning and executing specific data editing operations on Excel files.

## Proposed Solution

To enable the AI agent to understand the content of Excel files, we have developed a Python script that converts each sheet of a `.xlsx` file into a CSV file, which is readable by both humans and AI.

### Expected Workflow

1. A user places a `.xlsx` file in the designated `input` directory.
2. A user runs the Python script (`xlsx_to_csv.py`).
3. The script reads the `.xlsx` file, generates a separate CSV file for each sheet, and outputs them to the newly created `csv_output` directory.
4. The AI agent reads the generated CSV files to indirectly understand the structure and content of the Excel file.
5. In the future, based on this understanding, the AI agent will write code in `custom_operation.py` to manipulate the Excel file according to the user's instructions.

## Your Role

Your primary responsibility is to **edit `custom_operation.py`** based on the user's request.

- **Analyze the CSV files** in the `csv_output` directory to understand the structure and content of the Excel file.
- **Identify the target cells** for the operation and determine the appropriate values or formulas to insert.
- **Write Python code** in `custom_operation.py` to perform the requested Excel operations.
  - Use the `openpyxl` library to read, modify, and write `.xlsx` files.
  - The relationship between the CSV data and the Excel sheet is as follows: the first row and column in the CSV correspond to the Excel sheet structure. Use this to calculate the correct cell references.

## Important Notes

- The CSV files are a **read-only representation** of the Excel file. Any changes you make to the CSV files will not be reflected in the Excel file.
- Your code in `custom_operation.py` should **directly operate on the `.xlsx` file** in the `input` directory.
- When inserting formulas, you need to **infer the correct formula** based on the context of the surrounding cells. The CSV file will only show the calculated values.
- The current scope of this task is limited to the implementation of `xlsx_to_csv.py`. The logic for editing `custom_operation.py` will be handled in a future task.
