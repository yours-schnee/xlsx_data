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
    print(f"Found {len(excel_files)} Excel files")
    data_dict = {}
    
    for excel_file in excel_files:
        print(f"\nProcessing file: {excel_file}")
        df = pd.read_excel(excel_file)
        print(f"Columns found: {df.columns.tolist()}")
        
        for col in df.columns:
            print(f"\nProcessing column: {col}")
            # 各列の内容を表示
            print(f"Column data:\n{df[col].head()}")
            
            # 空でない値のみを取得し、内容を表示
            col_data = df[col].dropna()
            print(f"Non-empty values:\n{col_data.head()}")
            
            if len(col_data) >= 1:
                sheet_name = clean_sheet_name(str(col_data.iloc[0]))
                values = col_data.iloc[1:].tolist()
                
                print(f"Sheet name: {sheet_name}")
                print(f"Values: {values[:5]}...")  # 最初の5つの値を表示
                
                if sheet_name not in data_dict:
                    data_dict[sheet_name] = []
                data_dict[sheet_name].append([excel_file.name, values])
    
    print("\nCollected data summary:")
    for sheet_name, data in data_dict.items():
        print(f"Sheet: {sheet_name}, Files: {len(data)}")
    
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
    print(f"\nSaved to {output_file}")

if __name__ == "__main__":
    input_directory = "excel_files"
    output_file = "combined_output.xlsx"
    process_multiple_excel(input_directory, output_file)
