import pandas as pd
import os
import sys

def convert_excel_to_csv(input_dir, output_dir):
    """
    Read all Excel files in the specified directory and convert each sheet
    to CSV files saved in the output directory.
    """
    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Find files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_dir, filename)

            try:
                # Read the Excel file
                xls = pd.ExcelFile(file_path, engine='openpyxl')

                # Process each sheet
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    # Convert DataFrame to CSV
                    csv_content = df.to_csv(index=False)

                    # Save as a CSV file
                    output_filename = f"{sheet_name}.csv"
                    output_path = os.path.join(output_dir, output_filename)

                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(csv_content)

                    print(f"Converted sheet '{sheet_name}' to '{output_path}'")

            except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
                print(f"Error reading Excel file '{filename}': {e}")
            except Exception as e:
                print(f"An unexpected error occurred while processing file '{filename}': {e}")

if __name__ == '__main__':
    # Specify directory paths
    input_directory = sys.argv[1] if len(sys.argv) > 1 else 'input'
    output_directory = sys.argv[2] if len(sys.argv) > 2 else 'csv_output'

    # Run the conversion process
    convert_excel_to_csv(input_directory, output_directory)
