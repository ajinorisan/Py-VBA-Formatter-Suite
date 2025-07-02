import sys
import time
import os
import hashlib
import psutil
import win32com.client
import win32gui
import win32process
import tkinter as tk
from tkinter import messagebox
import re
import pythoncom

# --- ▼ 自作VBAフォーマッター（変更なし・完成版）▼ ---
class VbaFormatter:
    def __init__(self, indent_char: str = "    "):
        self.indent_char = indent_char
        self.INDENT_KEYWORDS = ("if", "for", "do", "with", "sub", "public sub", "private sub", "function", "public function", "private function", "property", "public property", "private property", "select case", "type")
        self.DEDENT_KEYWORDS = ("end if", "next", "loop", "end with", "end sub", "end function", "end property", "end select", "end type")
        self.ELSE_KEYWORDS = ("else", "elseif", "else if")

    def _get_judgement_line(self, code_line: str) -> str:
        """文字列リテラルとコメントを除いた、キーワード判定用の行を返す"""
        clean_line = ""; in_string = False
        for char in code_line:
            if char == '"': in_string = not in_string; continue
            if char == "'" and not in_string: break
            if not in_string: clean_line += char
        return clean_line.strip()

    def format_code(self, code_string: str) -> str:
        """【最終完成版】あなたのアルゴリズムを忠実に再現した、究極のフォーマッター"""
        lines = code_string.splitlines(); formatted_lines = []; current_indent_level = 0
        was_else_or_elseif = False # 直前の行がElse/ElseIfだったかを記憶するフラグ

        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                if formatted_lines and formatted_lines[-1] != "":
                    formatted_lines.append("")
                continue

            judgement_line = self._get_judgement_line(stripped_line.replace("_", "")).lower()

            # 1. キーワードフラグの判定
            judgement_parts = judgement_line.split()
            first_word = judgement_parts[0] if judgement_parts else ""
            first_two_words = " ".join(judgement_parts[:2]) if len(judgement_parts) > 1 else ""
            
            is_indent = first_two_words in self.INDENT_KEYWORDS or first_word in self.INDENT_KEYWORDS
            is_dedent = first_two_words in self.DEDENT_KEYWORDS or first_word in self.DEDENT_KEYWORDS
            is_else = first_two_words in self.ELSE_KEYWORDS or first_word in self.ELSE_KEYWORDS

            # 2. インデントレベルの調整（デデント先行）
            if is_else:
                current_indent_level = max(0, current_indent_level - 1)
            elif is_dedent and not was_else_or_elseif:
                current_indent_level = max(0, current_indent_level - 1)

            # 3. 行の出力
            formatted_lines.append(self.indent_char * current_indent_level + stripped_line)

            # 4. インデントレベルの事後調整
            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then")
                rest_of_line = judgement_line[then_pos + 4:].strip()
                if rest_of_line and not rest_of_line.startswith("'"):
                    is_single_line_if = True

            if (is_indent and not is_single_line_if) or is_else:
                current_indent_level += 1
            
            # 5. 最後に、次のループのためにフラグを更新
            was_else_or_elseif = is_else
                
        return "\n".join(formatted_lines)
# --- ▲ 自作VBAフォーマッター ▲ ---

def get_excel_process_count():
    count = 0
    for proc in psutil.process_iter(['name']):
        if proc.info['name'].lower() == 'excel.exe': count += 1
    return count

def ask_to_exit():
    root = tk.Tk(); root.withdraw()
    response = messagebox.askyesno("VBA Formatter", "全てのExcelが終了しました。ツールを終了しますか？")
    root.destroy(); return response

def get_active_excel_info():
    """現在アクティブなExcelの (App, Hwnd, FilePath) を返す"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if psutil.Process(pid).name().lower() == 'excel.exe':
            app = win32com.client.GetActiveObject("Excel.Application")
            if app and app.ActiveWorkbook and app.ActiveWorkbook.FullName:
                return app, hwnd, app.ActiveWorkbook.FullName
    except: pass
    return None, None, None

def main():
    print("VBAフォーマッターを起動しました。 (Ctrl+Cで手動終了)")
    
    formatter = VbaFormatter() 
    target_app = None; target_hwnd = None; target_filepath = None
    last_modified_time = 0
    vbe_is_open = False

    exit_prompted = False

    try:
        while True:
            # 1. アクティブなExcelの情報を取得
            active_app, active_hwnd, active_filepath = get_active_excel_info()

            # 2. 監視対象の切り替え
            if active_app and active_hwnd != target_hwnd:
                if target_app:
                    target_app = None
                
                print(f"\n監視対象を切り替えました -> {os.path.basename(active_filepath)}")
                target_app = active_app
                target_hwnd = active_hwnd
                target_filepath = active_filepath
                vbe_is_open = False
                if os.path.exists(target_filepath):
                    last_modified_time = os.path.getmtime(target_filepath)
            
            # 3. 監視とフォーマットの実行
            if target_app and target_filepath:
                try:
                    # 3-a. 生存確認
                    if not win32gui.IsWindow(target_hwnd):
                        raise Exception("ウィンドウが存在しません")

                    # 3-b. VBEの開閉状態を安全にチェック
                    try:
                        vbe_is_open = target_app.VBE.MainWindow.Visible
                    except Exception:
                        vbe_is_open = False
                    
                    # --- ▼★ ここが、あなたの指摘を反映した最終修正箇所です ★▼ ---
                    if vbe_is_open:
                        # VBEが開かれている場合の処理
                        current_modified_time = os.path.getmtime(target_filepath)
                        if current_modified_time > last_modified_time:
                            print(f"\n[{os.path.basename(target_filepath)}] の保存を検知。")
                            last_modified_time = current_modified_time
                            
                            print(" -> フォーマットを実行します...")
                            try:
                                vbe = target_app.VBE
                                for component in vbe.ActiveVBProject.VBComponents:
                                    if component.CodeModule.CountOfLines > 0:
                                        original_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                                        formatted_code = formatter.format_code(original_code)
                                        if original_code != formatted_code:
                                            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                                            component.CodeModule.AddFromString(formatted_code)
                                print(" -> フォーマットが完了しました。")
                            except Exception as format_error:
                                print(f" -> [エラー] フォーマット中にエラーが発生: {format_error}")
                        
                        print(f"'{os.path.basename(target_filepath)}' のVBEを監視中...", end="\r")

                    else:
                        # VBEが閉じられている場合の処理
                        print("VBEが開かれるのを待っています...", end="\r")
                    # --- ▲★ continue を削除し、ループが必ず最後まで到達するように修正 ▲★ ---

                except Exception as e:
                    # <--- 変更点: ここが最も重要な修正箇所です ---
                    # Excelが閉じた、またはCOMエラーが発生した場合の処理
                    print(f"\n'{os.path.basename(target_filepath)}' の監視を中断しました。理由: {e}")
                    
                    # 1. 監視対象の情報をすべてリセット
                    target_app = None
                    target_hwnd = None
                    target_filepath = None
                    
                    # 2. ガベージコレクションを明示的に呼び出し、COM参照を確実に解放する
                    import gc; gc.collect()
                    print(" -> 接続をリセットし、再スキャンします。")
                    # この後、ループの先頭に戻り、別のExcelを探すか、終了処理に入る

            else: # 3. 監視対象がいない場合
                # <--- 変更点: 終了判定ロジックをここに集約 ---
                if get_excel_process_count() == 0:
                    if not exit_prompted:
                        print("\n全てのExcelが終了したようです。")
                        if ask_to_exit():
                            break # whileループを抜けてプログラムを終了
                        else:
                            exit_prompted = True # 「いいえ」が押されたら、次回Excel起動まで再質問しない
                    print("Excelの起動を待っています...", end="\r")
                else:
                    # Excelプロセスはあるが、アクティブなブックが見つからない状態
                    print("アクティブなExcelブックを探しています...", end="\r")
                    exit_prompted = False # Excelが起動しているので、必要なら再度終了確認できるようにフラグを戻す

            time.sleep(1)
        
    except KeyboardInterrupt:
        print("\nプログラムを終了します。")
    finally:
        if target_app:
            target_app = None
        import gc; gc.collect()
        print("リソースを解放しました。")


if __name__ == "__main__":
    main()
