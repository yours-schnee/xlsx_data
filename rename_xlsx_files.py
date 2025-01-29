import os

def rename_xlsx_files(root_path):
    # サブフォルダを取得
    for folder_name in os.listdir(root_path):
        folder_path = os.path.join(root_path, folder_name)
        
        # フォルダであることを確認
        if os.path.isdir(folder_path):
            # フォルダ内のファイルを取得
            for file_name in os.listdir(folder_path):
                # xlsxファイルを探す
                if file_name.endswith('.xlsx'):
                    old_file_path = os.path.join(folder_path, file_name)
                    new_file_name = f"{folder_name}.xlsx"
                    new_file_path = os.path.join(folder_path, new_file_name)
                    
                    # ファイル名を変更
                    os.rename(old_file_path, new_file_path)
                    print(f"Renamed: {file_name} -> {new_file_name}")

# 使用例
root_directory = "path/to/your/root/folder"
rename_xlsx_files(root_directory)