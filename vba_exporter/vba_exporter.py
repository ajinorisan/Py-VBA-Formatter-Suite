# ver 1.0.0 GitHub管理用にVBAをファイルに出力
# ver 1.0.1 フォーマッター機能追加
# ver 1.0.2 (フォーマッター更新)

import os
import win32com.client
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import re

OUTPUT_BASE_FOLDER = "vba_source"
VB_COMPONENT_TYPE = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}


# --- ▼ [手順1] 完成したVbaFormatterクラスをここに追加 ▼ ---
class VbaFormatter:
    """VBAコードのインデントを自動整形するクラス。"""
    def __init__(self, indent_char: str = "    "):
        self.indent_char = indent_char
        self.INDENT_KEYWORDS = (
            "if", "for", "do", "with", "sub", "public sub", "private sub", 
            "function", "public function", "private function", 
            "property", "public property", "private property", 
            "select case", "type"
        )
        self.DEDENT_KEYWORDS = (
            "end if", "next", "loop", "end with", "end sub", 
            "end function", "end property", "end select", "end type"
        )
        self.MID_BLOCK_KEYWORDS = (
            "else", "elseif", "else if"
        )

    def _get_judgement_line(self, code_line: str) -> str:
        """文字列リテラルとコメントを除いた、キーワード判定用の行を返す。"""
        clean_line = ""; in_string = False
        for char in code_line:
            if char == '"': in_string = not in_string; continue
            if char == "'" and not in_string: break
            if not in_string: clean_line += char
        return clean_line.strip()

    def format_code(self, code_string: str) -> str:
        """VBAコード文字列を受け取り、整形後のコード文字列を返す。"""
        lines = code_string.splitlines()
        formatted_lines = []
        current_indent_level = 0
        block_stack = []

        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                if formatted_lines and formatted_lines[-1] != "":
                    formatted_lines.append("")
                continue

            judgement_line = self._get_judgement_line(stripped_line.replace("_", "")).lower()
            judgement_parts = judgement_line.split()
            first_word = judgement_parts[0] if judgement_parts else ""
            first_two_words = " ".join(judgement_parts[:2]) if len(judgement_parts) > 1 else ""

            is_start_block = first_two_words in self.INDENT_KEYWORDS or first_word in self.INDENT_KEYWORDS
            is_end_block = first_two_words in self.DEDENT_KEYWORDS or first_word in self.DEDENT_KEYWORDS
            is_mid_block = first_two_words in self.MID_BLOCK_KEYWORDS or first_word in self.MID_BLOCK_KEYWORDS
            is_case_statement = first_word == "case" or first_two_words == "case else"
            is_select_case = first_two_words == "select case"
            is_end_select = first_two_words == "end select"

            if is_end_select:
                current_indent_level = max(0, current_indent_level - 2)
                if block_stack: block_stack.pop()
            elif is_case_statement:
                if block_stack and block_stack[-1] == 'in_case':
                    current_indent_level = max(0, current_indent_level - 1)
            elif is_mid_block:
                current_indent_level = max(0, current_indent_level - 1)
            elif is_end_block:
                current_indent_level = max(0, current_indent_level - 1)
                if block_stack: block_stack.pop()

            formatted_lines.append(self.indent_char * current_indent_level + stripped_line)

            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then")
                rest_of_line = judgement_line[then_pos + 4:].strip()
                if rest_of_line and not rest_of_line.startswith("'"):
                    is_single_line_if = True

            if is_select_case:
                current_indent_level += 1
                block_stack.append('select')
            elif is_case_statement:
                current_indent_level += 1
                if block_stack and block_stack[-1] == 'select':
                    block_stack[-1] = 'in_case'
            elif (is_start_block and not is_single_line_if) or is_mid_block:
                current_indent_level += 1
                if is_start_block and not is_single_line_if:
                    block_stack.append('other')

        return "\n".join(formatted_lines)


class VbaExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("VBA Exporter (VBA Logic)")
        self.root.geometry("700x500")

        # --- ▼ [手順3] VbaFormatterをインスタンス化 ▼ ---
        self.formatter = VbaFormatter()

        self.log_area = scrolledtext.ScrolledText(
            root, wrap=tk.WORD, font=("Meiryo UI", 9)
        )
        self.log_area.pack(expand=True, fill="both", padx=10, pady=5)

        self.run_button = tk.Button(
            root, text="Excelファイルを選択して実行", command=self.start_export_thread
        )
        self.run_button.pack(pady=10)

        sys.stdout = self.RedirectText(self.log_area)
        sys.stderr = self.RedirectText(self.log_area)
    
    # --- ▼ [手順2] 古いフォーマット関連メソッドを削除 ▼ ---
    # _get_judgement_line と format_vba_indent はここから削除されました。
    
    def start_export_thread(self):
        """処理を別スレッドで開始する"""
        self.run_button.config(state=tk.DISABLED)
        self.log_area.delete("1.0", tk.END)

        thread = threading.Thread(target=self.run_export_process)
        thread.daemon = True
        thread.start()

    def run_export_process(self):
        """メインのエクスポート処理"""
        print("ファイル選択ダイアログを開きます...")
        selected_files = self.select_files()

        if not selected_files:
            print("ファイルが選択されなかったため、処理を中断しました。")
            self.run_button.config(state=tk.NORMAL)
            return

        print("VBAエクスポート処理を開始します...")

        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        output_dir = os.path.join(base_dir, OUTPUT_BASE_FOLDER)

        all_success = True
        for excel_filepath in selected_files:
            if not os.path.isfile(excel_filepath):
                print(f"[警告] 指定されたファイルが見つかりません: {excel_filepath}")
                all_success = False
                continue

            print(f"\nファイルを処理中: {excel_filepath}")

            excel_filename = os.path.basename(excel_filepath)
            output_folder_name = os.path.splitext(excel_filename)[0]
            output_folder_path = os.path.join(output_dir, output_folder_name)

            if not self.export_vba_from_file(excel_filepath, output_folder_path):
                all_success = False

        print("\nすべての処理が完了しました。")
        if not all_success:
            print("いくつかのファイルでエラーが発生しました。詳細は上記のログを確認してください。")

        self.run_button.config(state=tk.NORMAL)

    def select_files(self):
        """ファイル選択ダイアログを表示する"""
        file_paths = filedialog.askopenfilenames(
            title="VBAを抽出したいExcelファイルを選択してください (複数選択可)",
            filetypes=[("Excel Files", "*.xlsm *.xlsb *.xls")],
        )
        return file_paths

    def export_vba_from_file(self, excel_filepath, output_folder):
        """指定されたExcelファイルからVBAコードをエクスポートする"""
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
                    original_code = component.CodeModule.Lines(
                        1, component.CodeModule.CountOfLines
                    )
                    
                    try:
                        #print(f"    - Formatting {component.Name}...")
                        # --- ▼ [手順4] 新しいフォーマッターを呼び出す ▼ ---
                        formatted_code = self.formatter.format_code(original_code)
                        #print(f"    - {component.Name} のインデント整形完了")
                    except Exception as e:
                        print(f"    - [警告] {component.Name} のインデント整形に失敗: {e}")
                        formatted_code = original_code

                    with open(output_filepath, "w", encoding="utf-8") as f:
                        f.write(formatted_code)

            workbook.Close(SaveChanges=False)
            print(f"  [完了] {os.path.basename(excel_filepath)}")
            return True
        except Exception as e:
            print(f"  [エラー] {os.path.basename(excel_filepath)} の処理中にエラーが発生: {e}")
            return False
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception as e:
                    print(f"  [警告] Excelの終了処理中にエラーが発生しました: {e}", file=sys.stderr)

    # --- ▼ [手順5] 重複していたRedirectTextを削除し、1つに整理 ▼ ---
    class RedirectText:
        """printの出力をTextウィジェットにリダイレクトする"""
        def __init__(self, text_widget):
            self.output = text_widget

        def write(self, string):
            self.output.insert(tk.END, string)
            self.output.see(tk.END)

        def flush(self):
            pass


if __name__ == "__main__":
    root = tk.Tk()
    app = VbaExporterApp(root)
    root.mainloop()