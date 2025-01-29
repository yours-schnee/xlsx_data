import pandas as pd
import openpyxl
from openpyxl import Workbook
from pathlib import Path

def process_multiple_excel(input_dir, output_file):
    # 入力ディレクトリ内のすべてのxlsxファイルを取得
    excel_files = sorted(list(Path(input_dir).glob("*.xlsx")))
    
    # データを格納する辞書（キー：列名、値：[ファイル名, データのリスト]のリスト）
    data_dict = {}
    
    # 各Excelファイルを処理
    for excel_file in excel_files:
        df = pd.read_excel(excel_file)
        
        # 各列を処理
        for col in df.columns:
            # 空でない値を持つ行を取得
            col_data = df[col][df[col].notna()]
            
            # 1行目（シート名）を取得
            sheet_name = str(col_data.iloc[0])
            
            # 2行目以降のデータを取得
            values = col_data.iloc[1:].tolist()
            
            # データを辞書に追加
            if sheet_name not in data_dict:
                data_dict[sheet_name] = []
            # ファイル名とデータを組にして保存
            data_dict[sheet_name].append([excel_file.name, values])
    
    # 新しいワークブックを作成
    wb = Workbook()
    wb.remove(wb.active)  # デフォルトシートを削除
    
    # 各シートにデータを書き込む
    for sheet_name, file_data_list in data_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        
        # 各ファイルのデータを別々の列に書き込む
        for col_idx, (filename, values) in enumerate(file_data_list, start=1):
            # ファイル名を1行目に書き込む
            ws.cell(row=1, column=col_idx, value=filename)
            
            # データを2行目以降に書き込む
            for row_idx, value in enumerate(values, start=2):
                ws.cell(row=row_idx, column=col_idx, value=value)
    
    # 結果を保存
    wb.save(output_file)

# 使用例
if __name__ == "__main__":
    input_directory = "excel_files"
    output_file = "combined_output.xlsx"
    process_multiple_excel(input_directory, output_file)
