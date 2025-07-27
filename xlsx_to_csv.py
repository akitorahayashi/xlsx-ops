import pandas as pd
import os
import sys
import re

def convert_excel_to_csv(input_dir, output_dir):
    """
    Read all Excel files in the specified directory and convert each sheet
    to CSV files saved in the output directory.
    """
    # Validate input directory exists
    if not os.path.isdir(input_dir):
        print(f"Error: Input directory '{input_dir}' does not exist")
        return

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Find files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_dir, filename)

            try:
                # Read the Excel file
                with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                    # Process each sheet
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)

                        # Sanitize sheet name for file system
                        safe_sheet_name = re.sub(r'[<>:"/\\|?*]', '_', sheet_name)

                        # Convert DataFrame to CSV
                        csv_content = df.to_csv(index=False)

                        # Save as a CSV file
                        output_filename = f"{safe_sheet_name}.csv"
                        output_path = os.path.join(output_dir, output_filename)

                        try:
                            with open(output_path, 'w', encoding='utf-8') as f:
                                f.write(csv_content)
                        except (IOError, OSError) as e:
                            print(f"Error writing CSV file '{output_path}': {e}")
                            continue

                        print(f"Converted sheet '{sheet_name}' to '{output_path}'")

            except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
                print(f"Error reading Excel file '{filename}': {e}")
            except Exception as e:
                print(f"An unexpected error occurred while processing file '{filename}': {e}")

if __name__ == '__main__':
    # Specify directory paths
    input_directory = os.path.normpath(sys.argv[1]) if len(sys.argv) > 1 else 'input'
    output_directory = os.path.normpath(sys.argv[2]) if len(sys.argv) > 2 else 'csv_output'

    # Validate paths don't contain traversal attempts
    if '..' in input_directory or '..' in output_directory:
        print("Error: Directory paths cannot contain '..' for security reasons")
        sys.exit(1)

    # Run the conversion process
    convert_excel_to_csv(input_directory, output_directory)
