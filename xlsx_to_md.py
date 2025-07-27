import pandas as pd
import os

def convert_excel_to_markdown(input_dir, output_dir):
    """
    指定されたディレクトリ内のすべてのExcelファイルを読み込み、
    各シートをMarkdownファイルに変換して出力ディレクトリに保存する。
    """
    # 出力ディレクトリが存在しない場合は作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 入力ディレクトリ内のファイルを検索
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_dir, filename)

            try:
                # Excelファイルを読み込む
                xls = pd.ExcelFile(file_path)

                # 各シートを処理
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)

                    # DataFrameをCSVに変換
                    csv_content = df.to_csv(index=False)

                    # CSVファイルとして保存
                    output_filename = f"{sheet_name}.csv"
                    output_path = os.path.join(output_dir, output_filename)

                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(csv_content)


                    print(f"'{sheet_name}'シートを'{output_path}'に変換しました。")

            except Exception as e:
                print(f"ファイル '{filename}' の処理中にエラーが発生しました: {e}")

if __name__ == '__main__':
    # ディレクトリのパスを指定
    input_directory = 'input'
    output_directory = 'md_output'

    # 変換処理を実行
    convert_excel_to_markdown(input_directory, output_directory)
