# Excel to Markdown Converter

This project provides a Python script to convert Excel files (.xlsx) into Markdown files, making them readable and understandable for AI agents.

## Overview

The primary goal of this project is to enable AI agents to autonomously interact with Excel files. Since AI agents cannot directly read the binary content of `.xlsx` files, this script converts each sheet of an Excel file into a separate Markdown file. This allows the agent to understand the structure and content of the Excel file, enabling it to perform data manipulation tasks.

## Directory Structure

```
/
|-- xlsx_to_md.py
|-- md_output/
|   |-- sheet1.md
|   `-- product_list.md
|-- input/
|   `-- example.xlsx
|-- custom_operation.py
|-- main.py
...
```

- **`xlsx_to_md.py`**: The Python script that converts Excel files to Markdown.
- **`md_output/`**: The directory where the output Markdown files are saved.
- **`input/`**: The directory where you should place your `.xlsx` files.
- **`custom_operation.py`**: A script where the AI agent will write code to manipulate the Excel file based on its understanding of the Markdown files.
- **`main.py`**: The main script to run the entire process.

## How to Use

1. **Place your Excel file in the `input` directory.**
   - Make sure the file has a `.xlsx` extension.

2. **Run the conversion script.**
   ```bash
   python xlsx_to_md.py
   ```

3. **Check the `md_output` directory.**
   - The script will generate a Markdown file for each sheet in the Excel file. The filename will be `{sheet_name}.md`.

## Markdown Format

The script uses the `pandas` library to convert each sheet into a Markdown table.

- **Formulas**: Formulas in the Excel file will be output as their calculated values.
- **Cell Positioning**: The top-left header of the Markdown table corresponds to cell A1 in the Excel sheet. The AI agent uses this correspondence to interpret the cell positions.

Example output:

|    | Product | Price | Sales |
|---:|:---|---:|---:|
| 0 | Apple | 150 | 30 |
| 1 | Orange | 100 | 50 |
| 2 | Total | | 80 |

## Expected Benefits

- **Data Structure Comprehension**: The AI agent can intuitively understand the data structure of each sheet.
- **Rational Decision Making**: By understanding the context of the data, the AI agent can perform reasonable operations based on abstract instructions.
