import pandas as pd
import openpyxl
from openpyxl import Workbook
from pathlib import Path

def clean_sheet_name(name):
    invalid_chars = [':', '/', '\\', '?', '*', '[', ']']
    clean_name = str(name)
    for char in invalid_chars:
        clean_name = clean_name.replace(char, '_')
    return clean_name[:31]

def process_multiple_excel(input_dir, output_file):
    excel_files = sorted(list(Path(input_dir).glob("*.xlsx")))
    data_dict = {}
    
    # データの読み込みと整理
    for excel_file in excel_files:
        print(f"Processing file: {excel_file}")
        df = pd.read_excel(excel_file)
        
        for col in df.columns:
            # 空でない値のみを取得
            col_data = df[col].dropna()
            
            if len(col_data) >= 1:
                # 最初の行をシート名として使用
                sheet_name = clean_sheet_name(str(col_data.iloc[0]))
                # 2行目以降をデータとして使用
                values = col_data.iloc[1:].tolist()
                
                if sheet_name not in data_dict:
                    data_dict[sheet_name] = []
                data_dict[sheet_name].append([excel_file.name, values])
                print(f"Found data for sheet: {sheet_name}")
    
    # 新しいワークブックの作成
    wb = Workbook()
    default_sheet = wb.active
    default_sheet.title = "Sheet1"
    
    # データの書き込み
    for sheet_name, file_data_list in data_dict.items():
        print(f"Creating sheet: {sheet_name}")
        ws = wb.create_sheet(title=sheet_name)
        
        for col_idx, (filename, values) in enumerate(file_data_list, start=1):
            # ファイル名をヘッダーとして書き込み
            ws.cell(row=1, column=col_idx, value=filename)
            print(f"Writing data from {filename}")
            
            # データの書き込み
            for row_idx, value in enumerate(values, start=2):
                if value is not None and value != "":
                    ws.cell(row=row_idx, column=col_idx, value=value)
    
    # デフォルトシートの処理
    if len(data_dict) > 0:
        wb.remove(default_sheet)
    
    # ファイルの保存
    wb.save(output_file)
    print(f"Successfully saved data to {output_file}")

if __name__ == "__main__":
    input_directory = "excel_files"
    output_file = "combined_output.xlsx"
    process_multiple_excel(input_directory, output_file)
