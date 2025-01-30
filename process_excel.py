import pandas as pd
import openpyxl
from openpyxl import Workbook
from pathlib import Path

def clean_sheet_name(name):
    # Excelのシート名として無効な文字を除去
    invalid_chars = [':', '/', '\\', '?', '*', '[', ']']
    clean_name = str(name)
    for char in invalid_chars:
        clean_name = clean_name.replace(char, '_')
    # シート名の長さを31文字に制限（Excelの制限）
    return clean_name[:31]

def process_multiple_excel(input_dir, output_file):
    excel_files = sorted(list(Path(input_dir).glob("*.xlsx")))
    data_dict = {}
    
    for excel_file in excel_files:
        try:
            df = pd.read_excel(excel_file)
            
            for col in df.columns:
                col_data = df[col][df[col].notna()]
                
                if len(col_data) > 0:  # データが存在する場合のみ処理
                    sheet_name = clean_sheet_name(str(col_data.iloc[0]))
                    values = col_data.iloc[1:].tolist()
                    
                    if sheet_name not in data_dict:
                        data_dict[sheet_name] = []
                    data_dict[sheet_name].append([excel_file.name, values])
        except Exception as e:
            print(f"Error processing {excel_file}: {str(e)}")
    
    wb = Workbook()
    wb.remove(wb.active)
    
    for sheet_name, file_data_list in data_dict.items():
        try:
            ws = wb.create_sheet(title=sheet_name)
            
            for col_idx, (filename, values) in enumerate(file_data_list, start=1):
                ws.cell(row=1, column=col_idx, value=filename)
                
                for row_idx, value in enumerate(values, start=2):
                    # データ型の検証と変換
                    if pd.isna(value):
                        continue
                    ws.cell(row=row_idx, column=col_idx, value=str(value))
        except Exception as e:
            print(f"Error creating sheet {sheet_name}: {str(e)}")
    
    try:
        wb.save(output_file)
        print(f"Successfully saved to {output_file}")
    except Exception as e:
        print(f"Error saving workbook: {str(e)}")

if __name__ == "__main__":
    input_directory = "excel_files"
    output_file = "combined_output.xlsx"
    process_multiple_excel(input_directory, output_file)
