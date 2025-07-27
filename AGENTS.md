# AI Agent Development Guidelines

This document's purpose is to ensure that you, the AI agent, clearly understand the rationale behind this repository, the role of each file, and your specific responsibilities. It should not contain user-facing instructions like "prepare the .xlsx file" or "run the xlsx_to_csv.py script." This document outlines your workflow only.

## Your primary role is to act as an autonomous agent that modifies an Excel file based on user requests. Here is your workflow:

Analyze the Data Structure: Begin by reading the CSV files located in the csv_output directory. These files are generated from the user's original .xlsx file and serve as your map to understand its structure, sheet names, and data layout.

Generate Python Code: Based on the user's instructions and your analysis of the CSV files, write Python code in the custom_operation.py script. This code will use the `openpyxl` and `pandas` libraries to perform the required edits on the original .xlsx file.

Adhere to Constraints: Your only task is to write code within the custom_operation.py file. You must not modify any other file, including main.py.

Goal: Your generated code will be executed via main.py. The final, modified .xlsx file should be saved in the output directory. Your task is complete when the correct output is generated.
