import glob
from custom_operation import custom_operation

def find_excel_file():
    """
    Finds an Excel file in the current directory and returns its name.
    """
    excel_files = glob.glob("*.xlsx")
    if len(excel_files) == 1:
        return excel_files[0]
    else:
        return None

if __name__ == "__main__":
    excel_file = find_excel_file()
    if excel_file:
        print(f"Excelファイルが見つかりました: '{excel_file}'")
        processed_file = custom_operation(excel_file)
        print("処理が完了しました。")
    else:
        print("Excelファイルが見つかりませんでした。")