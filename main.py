
import glob
import shutil
import os
from custom_operation import custom_operation

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
        print(f"Excelファイルが見つかりました: '{excel_file}'")
        processed_file = custom_operation(excel_file)
        output_file = os.path.join("output", os.path.basename(processed_file))
        shutil.copy(processed_file, output_file)
        print(f"処理が完了し、'{output_file}' に保存されました。")
    else:
        print("Excelファイルが見つかりませんでした。")
