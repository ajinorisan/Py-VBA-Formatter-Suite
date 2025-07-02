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
from pywintypes import com_error

'''# --- ▼ 自作VBAフォーマッター（変更なし・完成版）▼ ---
class VbaFormatter:
    """VBAコードのインデントを自動整形するクラス。"""
    def __init__(self, indent_char: str = "    "):
        self.indent_char = indent_char
        self.INDENT_KEYWORDS = ("if", "for", "do", "with", "sub", "public sub", "private sub", "function", "public function", "private function", "property", "public property", "private property", "select case", "type")
        self.DEDENT_KEYWORDS = ("end if", "next", "loop", "end with", "end sub", "end function", "end property", "end select", "end type")
        self.ELSE_KEYWORDS = ("else", "elseif", "else if")

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
        lines = code_string.splitlines(); formatted_lines = []; current_indent_level = 0
        was_else_or_elseif = False

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

            # 5. 次のループのためにフラグを更新
            was_else_or_elseif = is_else

        return "\n".join(formatted_lines)
# --- ▲ 自作VBAフォーマッター ▲ ---'''

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
        block_stack = []  # [修正点] ブロックの種類を管理するスタック

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

            # --- キーワード判定 ---
            is_start_block = first_two_words in self.INDENT_KEYWORDS or first_word in self.INDENT_KEYWORDS
            is_end_block = first_two_words in self.DEDENT_KEYWORDS or first_word in self.DEDENT_KEYWORDS
            is_mid_block = first_two_words in self.MID_BLOCK_KEYWORDS or first_word in self.MID_BLOCK_KEYWORDS
            is_case_statement = first_word == "case" or first_two_words == "case else"
            is_select_case = first_two_words == "select case"
            is_end_select = first_two_words == "end select"

            # --- デデント処理 (先行) ---
            if is_end_select:
                # [修正点] CaseブロックとSelectブロックの2段階デデント
                current_indent_level = max(0, current_indent_level - 2)
                if block_stack: block_stack.pop()
            elif is_case_statement:
                # [修正点] 2つ目以降のCaseの場合、前のCaseブロックを閉じる
                if block_stack and block_stack[-1] == 'in_case':
                    current_indent_level = max(0, current_indent_level - 1)
            elif is_mid_block:
                current_indent_level = max(0, current_indent_level - 1)
            elif is_end_block:
                current_indent_level = max(0, current_indent_level - 1)
                if block_stack: block_stack.pop()

            # --- 行の出力 ---
            formatted_lines.append(self.indent_char * current_indent_level + stripped_line)

            # --- インデント処理 (事後) ---
            is_single_line_if = False
            if first_word == "if" and "then" in judgement_line:
                then_pos = judgement_line.find("then")
                rest_of_line = judgement_line[then_pos + 4:].strip()
                if rest_of_line and not rest_of_line.startswith("'"):
                    is_single_line_if = True

            if is_select_case:
                current_indent_level += 1
                block_stack.append('select') # 'select'ブロック開始
            elif is_case_statement:
                current_indent_level += 1
                if block_stack and block_stack[-1] == 'select':
                    block_stack[-1] = 'in_case' # 'select'から'in_case'状態へ移行
            elif (is_start_block and not is_single_line_if) or is_mid_block:
                current_indent_level += 1
                if is_start_block and not is_single_line_if:
                    block_stack.append('other') # 'select'以外のブロック

        return "\n".join(formatted_lines)

def get_excel_process_count() -> int:
    """実行中の 'excel.exe' プロセスの数を返す。"""
    count = 0
    for proc in psutil.process_iter(['name']):
        if proc.info['name'].lower() == 'excel.exe':
            count += 1
    return count

def ask_to_exit() -> bool:
    """ツールを終了するか確認するダイアログを表示する。"""
    root = tk.Tk()
    root.withdraw() # メインウィンドウを非表示にする
    response = messagebox.askyesno("VBA Formatter", "全てのExcelが終了しました。ツールを終了しますか？")
    root.destroy()
    return response

def get_active_excel_info():
    """
    フォアグラウンドのExcelアプリケーション情報を取得する。
    
    Returns:
        tuple: (Excel.Application オブジェクト, ウィンドウハンドル, ファイルパス)
               アクティブなExcelが見つからない、または情報が取得できない場合は (None, None, None) を返す。
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        # ウィンドウハンドルが無効な場合はここで処理を中断
        if not hwnd:
            return None, None, None
            
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if psutil.Process(pid).name().lower() == 'excel.exe':
            app = win32com.client.GetActiveObject("Excel.Application")
            if app and getattr(app, 'ActiveWorkbook', None) and app.ActiveWorkbook.FullName:
                return app, hwnd, app.ActiveWorkbook.FullName
                
    except (psutil.NoSuchProcess, psutil.AccessDenied, com_error):
        # 発生しうる例外（プロセス非存在、アクセス拒否、COMエラー）を明示的に捕捉します。
        # これらはウィンドウ切替時などに正常な挙動として発生しうるため、エラーとせず処理を続行します。
        pass
    
    return None, None, None

def main():
    """メイン処理。Excelの監視とVBAコードの自動フォーマットを実行する。"""
    print("VBAフォーマッターを起動しました。 (Ctrl+Cで手動終了)")

    formatter = VbaFormatter()
    target_app = None
    target_hwnd = None
    target_filepath = None
    last_modified_time = 0

    # 直前のループでのExcelプロセスの存在状態を記録
    was_excel_running = False

    # 起動直後に実行中のExcelを誤認識しないための待機
    print("初期スキャン待機中...")
    time.sleep(2)

    try:
         # COMライブラリを現在のスレッドで初期化
        pythoncom.CoInitialize()
        while True:
            # 1. フォアグラウンドのExcel情報を取得
            active_app, active_hwnd, active_filepath = get_active_excel_info()

            # 2. 監視対象の切り替え処理
            # アクティブなExcelが変わり、かつそれが現在の監視対象と異なる場合に実行
            if active_app and active_hwnd != target_hwnd:
                print(f"\n監視対象を切り替えました -> {os.path.basename(active_filepath)}")
                target_app, target_hwnd, target_filepath = active_app, active_hwnd, active_filepath
                if os.path.exists(target_filepath):
                    last_modified_time = os.path.getmtime(target_filepath)

            # 3. 監視対象の喪失処理
            # アクティブウィンドウがExcelでない、または監視対象ウィンドウが閉じられた場合に実行
            elif active_hwnd is None or (target_hwnd and not win32gui.IsWindow(target_hwnd)):
                if target_hwnd:
                    print(f"\n'{os.path.basename(target_filepath)}' の監視を中断します。")
                target_app, target_hwnd, target_filepath = None, None, None
                # [注意] COMオブジェクトの参照を解放後、ガベージコレクションを明示的に呼び出すことで
                # リソースリークのリスクを低減できますが、必ずしも即時解放を保証するものではありません。
                import gc; gc.collect()

            # 4. フォーマット実行処理
            # 監視対象のExcelが存在する場合に実行
            if target_app:
                try:
                    # VBEウィンドウが開いている場合のみ処理を続行
                    if target_app.VBE.MainWindow.Visible:
                        current_modified_time = os.path.getmtime(target_filepath)
                        # ファイルの保存（更新時刻の変更）を検知
                        if current_modified_time > last_modified_time:
                            print(f"\n[{os.path.basename(target_filepath)}] の保存を検知し、フォーマットを実行します。")
                            last_modified_time = current_modified_time

                            vbe = target_app.VBE
                            for component in vbe.ActiveVBProject.VBComponents:
                                # [要確認] プロジェクトがパスワードで保護されている場合、component.CodeModuleへのアクセスでエラーが発生します。
                                # このループ内、またはアクセス直前にtry-exceptブロックを設けると、より堅牢になります。
                                if component.CodeModule.CountOfLines > 0:
                                    original_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
                                    formatted_code = formatter.format_code(original_code)
                                    # コードに変更があった場合のみ書き込みを実行
                                    if original_code != formatted_code:
                                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                                        component.CodeModule.AddFromString(formatted_code)
                            print(" -> フォーマットが完了しました。")

                        print(f"'{os.path.basename(target_filepath)}' のVBEを監視中...", end="\r")
                    else:
                        print(f"'{os.path.basename(target_filepath)}' のVBEが開かれるのを待っています...", end="\r")
                except Exception as e:
                    # COM接続エラーなど、監視対象との通信で問題が発生した場合
                    print(f"\n'{os.path.basename(target_filepath)}' との接続エラー: {e}")
                    target_app, target_hwnd, target_filepath = None, None, None
                    import gc; gc.collect()

            # 5. 終了判定ロジック
            # 監視対象のExcelがない場合に実行
            else:
                is_excel_running_now = get_excel_process_count() > 0

                # 直前のループではExcelが実行中だったが、現在は全て終了している場合
                if was_excel_running and not is_excel_running_now:
                    print("\n全てのExcelが終了したようです。")
                    if ask_to_exit():
                        break # メインループを抜けてプログラムを終了

                # 待機メッセージの表示
                if is_excel_running_now:
                    print("アクティブなExcelブックを探しています...", end="\r")
                else:
                    print("Excelの起動を待っています...", end="\r")

                # 現在のExcel実行状態を次回のループ用に保存
                was_excel_running = is_excel_running_now

            time.sleep(1)

    except KeyboardInterrupt:
        print("\nプログラムを終了します。")
    finally:
        # プログラム終了時にリソースをクリーンアップ
        if target_app:
            target_app = None

        # COMライブラリを解放
        pythoncom.CoUninitialize()

        import gc; gc.collect()
        print("リソースを解放しました。")


if __name__ == "__main__":
    main()
