import os
import win32com.client
import sys
import tkinter as tk
from tkinter import filedialog

# 抽出したVBAコードを保存するフォルダ名です。
OUTPUT_BASE_FOLDER = "vba_source"

VB_COMPONENT_TYPE = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}

def export_vba_from_file(excel_filepath, output_folder):
    """ 指定されたExcelファイルからVBAコードをエクスポートする """
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_filepath)
        
        print(f"  [処理中] {os.path.basename(excel_filepath)}")
        os.makedirs(output_folder, exist_ok=True)
        
        for component in workbook.VBProject.VBComponents:
            ext = VB_COMPONENT_TYPE.get(component.Type)
            if not ext:
                continue

            output_filename = f"{component.Name}{ext}"
            output_filepath = os.path.join(output_folder, output_filename)
            
            if component.CodeModule.CountOfLines > 0:
                with open(output_filepath, "w", encoding="utf-8") as f:
                    f.write(component.CodeModule.Lines(1, component.CodeModule.CountOfLines))

        workbook.Close(SaveChanges=False)
        print(f"  [完了] {os.path.basename(excel_filepath)}")
        return True

    except Exception as e:
        print(f"  [エラー] {os.path.basename(excel_filepath)} の処理中にエラーが発生: {e}", file=sys.stderr)
        return False
    finally:
        if excel:
            excel.Quit()

def select_files():
    """ ファイル選択ダイアログを表示し、選択された複数ファイルのパスをタプルで返す """
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="VBAを抽出したいExcelファイルを選択してください (複数選択可)",
        filetypes=[("Excel Files", "*.xlsm *.xlsb *.xls"), ("All Files", "*.*")]
    )
    root.destroy()
    return file_paths

def main():
    """ メイン処理 """
    print("ファイル選択ダイアログを開きます...")
    selected_files = select_files()

    if not selected_files:
        print("ファイルが選択されなかったため、処理を中断しました。")
        return

    print("VBAエクスポート処理を開始します...")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, OUTPUT_BASE_FOLDER)
    
    all_success = True
    for excel_filepath in selected_files:
        if not os.path.isfile(excel_filepath):
            print(f"[警告] 指定されたファイルが見つかりません: {excel_filepath}", file=sys.stderr)
            all_success = False
            continue

        print(f"\nファイルを処理中: {excel_filepath}")
        
        excel_filename = os.path.basename(excel_filepath)
        output_folder_name = os.path.splitext(excel_filename)[0]
        output_folder_path = os.path.join(output_dir, output_folder_name)

        if not export_vba_from_file(excel_filepath, output_folder_path):
            all_success = False
            
    print("\nすべての処理が完了しました。")
    if not all_success:
        print("いくつかのファイルでエラーが発生しました。詳細は上記のログを確認してください。")

if __name__ == "__main__":
    main()