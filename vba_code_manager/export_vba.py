# ver 1.0.1 (安定版)
# ・Excelが閉じられなかった時のエラーを修正
# ・exeファイルとして実行した際のパスの挙動を修正

import os
import win32com.client
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import re

OUTPUT_BASE_FOLDER = "vba_source"
VB_COMPONENT_TYPE = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}


class VbaExporterApp:
    # --- ▼★ ここからが新しいクラスの全貌です ★▼ ---

    def __init__(self, root):
        self.root = root
        self.root.title("VBA Exporter (VBA Logic)")
        self.root.geometry("700x500")

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

    def _get_judgement_line(self, code_line: str) -> str:
        """
        VBAの DoubleQuateEject と、コメント部分を落とすロジックを再現。
        文字列リテラルとコメントを除いた、キーワード判定用の行を返す。
        """
        clean_line = ""
        in_string = False
        for char in code_line:
            if char == '"':
                in_string = not in_string
                continue
            
            if char == "'" and not in_string:
                break
            
            if not in_string:
                clean_line += char
        
        return clean_line.strip()

    def format_vba_indent(self, code_string: str, indent_char: str = "    ") -> str:
        """【最終完成版】あなたのアルゴリズムを、Pythonのベストプラクティスで記述した最終形"""

        # --- ▼★ あなたのロジックを、洗練されたPythonコードで実装 ★▼ ---
        
        INDENT_KEYWORDS = ("if", "for", "do", "with", "sub", "public sub", "private sub", "function", "public function", "private function", "property", "public property", "private property", "select case", "type")
        # 'Else' はデデントとインデントの両方の性質を持つため、独立して扱う
        DEDENT_KEYWORDS = ("end if", "next", "loop", "end with", "end sub", "end function", "end property", "end select", "end type")
        ELSE_KEYWORDS = ("else", "elseif", "else if")

        lines = code_string.splitlines()
        formatted_lines = []
        current_indent_level = 0
        
        # 直前の行がElse/ElseIfだったかを記憶するフラグ
        was_else_or_elseif = False

        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                if formatted_lines and formatted_lines[-1] != "":
                    formatted_lines.append("")
                continue

            judgement_line = self._get_judgement_line(stripped_line.replace("_", "")).lower()

            # --- 1. キーワードフラグの判定 ---
            
            # 判定用の行から、最初の単語と2単語を取得
            judgement_parts = judgement_line.split()
            first_word = judgement_parts[0] if judgement_parts else ""
            first_two_words = " ".join(judgement_parts[:2]) if len(judgement_parts) > 1 else ""
            
            # 現在の行がどのキーワードに該当するかを判定
            is_indent = first_two_words in INDENT_KEYWORDS or first_word in INDENT_KEYWORDS
            is_dedent = first_two_words in DEDENT_KEYWORDS or first_word in DEDENT_KEYWORDS
            is_else = first_two_words in ELSE_KEYWORDS or first_word in ELSE_KEYWORDS

            # --- 2. インデントレベルの調整（デデント先行） ---
            
            # Else/ElseIfが来たら、まずデデント
            if is_else:
                current_indent_level = max(0, current_indent_level - 1)
            # End If などが来て、かつ直前がElseではなかったらデデント
            elif is_dedent and not was_else_or_elseif:
                current_indent_level = max(0, current_indent_level - 1)

            # --- 3. 行の出力 ---
            output_line = indent_char * current_indent_level + stripped_line
            formatted_lines.append(output_line)

            # --- 4. インデントレベルの事後調整 ---
            
            # 1行Ifはインデントしない
            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then")
                rest_of_line = judgement_line[then_pos + 4:].strip()
                if rest_of_line and not rest_of_line.startswith("'"):
                    is_single_line_if = True

            # Indentキーワード、またはElseキーワードなら、次の行からインデントを上げる
            if (is_indent and not is_single_line_if) or is_else:
                current_indent_level += 1
            
            # --- 5. 最後に、次のループのためにフラグを更新 ---
            was_else_or_elseif = is_else
                
        return "\n".join(formatted_lines)
        
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
                        print(f"    - Formatting {component.Name}...")
                        formatted_code = self.format_vba_indent(original_code)
                        print(f"    - {component.Name} のインデント整形完了")
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

    class RedirectText:
        """printの出力をTextウィジェットにリダイレクトする"""
        def __init__(self, text_widget):
            self.output = text_widget

        def write(self, string):
            self.output.insert(tk.END, string)
            self.output.see(tk.END)

        def flush(self):
            pass
    
    # --- ▲★ ここまでが新しいクラスの全貌です ★▼ ---

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
