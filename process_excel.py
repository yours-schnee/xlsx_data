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

def process_directory_excel(directory_path):
    excel_files = sorted(list(Path(directory_path).glob("*.xlsx")))
    if not excel_files:
        return None
    
    print(f"\nProcessing directory: {directory_path}")
    print(f"Found {len(excel_files)} Excel files")
    data_dict = {}
    
    for excel_file in excel_files:
        print(f"Processing file: {excel_file}")
        df = pd.read_excel(excel_file)
        
        for col in df.columns:
            col_data = df[col].dropna()
            
            if len(col_data) >= 1:
                sheet_name = clean_sheet_name(str(col_data.iloc[0]))
                values = col_data.iloc[1:].tolist()
                
                if sheet_name not in data_dict:
                    data_dict[sheet_name] = []
                data_dict[sheet_name].append([excel_file.name, values])
    
    return data_dict

def save_workbook(data_dict, output_file):
    if not data_dict:
        return
    
    wb = Workbook()
    default_sheet = wb.active
    default_sheet.title = "Sheet1"
    
    for sheet_name, file_data_list in data_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        
        for col_idx, (filename, values) in enumerate(file_data_list, start=1):
            ws.cell(row=1, column=col_idx, value=filename)
            
            for row_idx, value in enumerate(values, start=2):
                if value is not None and value != "":
                    ws.cell(row=row_idx, column=col_idx, value=value)
    
    if len(data_dict) > 0:
        wb.remove(default_sheet)
    
    wb.save(output_file)
    print(f"Saved to {output_file}")

def process_multiple_excel(input_dir):
    root_path = Path(input_dir)
    
    # サブディレクトリを取得
    subdirs = [d for d in root_path.iterdir() if d.is_dir()]
    print(f"Found {len(subdirs)} subdirectories")
    
    # 各サブディレクトリに対して処理を実行
    for subdir in subdirs:
        data_dict = process_directory_excel(subdir)
        if data_dict:
            # サブディレクトリ名を出力ファイル名に使用
            output_file = root_path / f"{subdir.name}_output.xlsx"
            save_workbook(data_dict, output_file)

if __name__ == "__main__":
    input_directory = "excel_files"
    process_multiple_excel(input_directory)
