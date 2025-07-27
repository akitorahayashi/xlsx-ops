
import glob
import shutil
import os
from custom_operation import custom_operation
from dotenv import load_dotenv
from openpyxl import load_workbook
def find_excel_file():
    """
    Finds an Excel file in the 'input' directory and returns its name.
    """
    excel_files = glob.glob("input/*.xlsx")
    if len(excel_files) == 1:
        return excel_files[0]
    else:
        return None

if __name__ == "__main__":
    excel_file = find_excel_file()
    if excel_file:
        print(f"Found Excel file: '{excel_file}'")
        processed_file = custom_operation(excel_file)

        if os.path.exists(".env"):
            load_dotenv()
            author_name = os.getenv("EXCEL_AUTHOR")
            if author_name:
                try:
                    workbook = load_workbook(processed_file)
                    properties = workbook.properties
                    properties.creator = author_name
                    properties.lastModifiedBy = author_name
                    workbook.save(processed_file)
                    print(f"Author set to: {author_name}")
                except Exception as e:
                    print(f"An error occurred while setting properties: {e}")

        output_file = os.path.join("output", os.path.basename(processed_file))
        shutil.copy(processed_file, output_file)
        print(f"Processing complete. Saved to '{output_file}'.")
    else:
        print("Excel file not found.")
