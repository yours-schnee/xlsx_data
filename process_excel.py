import pandas as pd
import openpyxl
from openpyxl import Workbook
from pathlib import Path

def process_multiple_excel(input_dir):
    root_path = Path(input_dir)
    
    # サブディレクトリを取得
    subdirs = [d for d in root_path.iterdir() if d.is_dir()]
    print(f"処理するサブディレクトリ数: {len(subdirs)}")
    
    # 各サブディレクトリを処理
    for subdir in subdirs:
        print(f"\n{subdir.name} を処理中...")
        
        # サブディレクトリ内のExcelファイルを取得
        excel_files = list(subdir.glob("*.xlsx"))
        if not excel_files:
            print(f"{subdir.name} にExcelファイルが見つかりません")
            continue
            
        print(f"見つかったExcelファイル数: {len(excel_files)}")
        
        # 新しいワークブックを作成
        wb = Workbook()
        wb.remove(wb.active)
        
        # 各Excelファイルを処理
        for excel_file in excel_files:
            print(f"{excel_file.name} を読み込み中...")
            df = pd.read_excel(excel_file)
            
            # 各列を処理
            for column in df.columns:
                # 空でない値のみを取得
                values = df[column].dropna().tolist()
                if len(values) < 2:  # シート名とデータの最小要件
                    continue
                
                # シート名を取得（1行目）
                sheet_name = str(values[0])[:31]  # Excel制限：31文字まで
                
                # シートが存在しない場合は作成
                if sheet_name not in wb.sheetnames:
                    wb.create_sheet(sheet_name)
                ws = wb[sheet_name]
                
                # 次の空き列を見つける
                next_col = ws.max_column + 1
                
                # ファイル名を1行目に書き込み
                ws.cell(row=1, column=next_col, value=excel_file.name)
                
                # 2行目以降にデータを書き込み
                for row_idx, value in enumerate(values[1:], start=2):
                    ws.cell(row=row_idx, column=next_col, value=value)
        
        # 出力ファイルを保存
        output_file = root_path / f"{subdir.name}_output.xlsx"
        wb.save(output_file)
        print(f"保存完了: {output_file}")

if __name__ == "__main__":
    input_directory = "excel_files"
    process_multiple_excel(input_directory)
