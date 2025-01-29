import pandas as pd
import openpyxl
from openpyxl import Workbook

def process_excel(input_file, output_file):
    # 入力ファイルを読み込む
    df = pd.read_excel(input_file)
    
    # 空でない値を持つ行を取得
    non_empty_rows = df.notna()
    
    # 新しいワークブックを作成
    wb = Workbook()
    wb.remove(wb.active)  # デフォルトシートを削除
    
    # 各列を処理
    for col in df.columns:
        # 列のデータを取得（空の値を除く）
        col_data = df[col][non_empty_rows[col]]
        
        # 1行目（シート名）を取得
        sheet_name = str(col_data.iloc[0])
        
        # 新しいシートを作成
        ws = wb.create_sheet(title=sheet_name)
        
        # 2行目以降のデータを書き込む
        for idx, value in enumerate(col_data.iloc[1:], start=1):
            ws.cell(row=idx, column=1, value=value)
    
    # 結果を保存
    wb.save(output_file)

# 使用例
if __name__ == "__main__":
    input_file = "input.xlsx"
    output_file = "output.xlsx"
    process_excel(input_file, output_file)