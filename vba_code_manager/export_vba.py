# ver 1.0.0 250624作成
# ver 1.0.1 エクセルが閉じられなかった時のエラー修正。exeにした時の挙動修正。
import os
import win32com.client
import sys
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading

OUTPUT_BASE_FOLDER = "vba_source"
VB_COMPONENT_TYPE = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}


class VbaExporterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("VBA Exporter")
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
            # .exeファイルとして実行されている場合のパス設定
            base_dir = os.path.dirname(sys.executable)
        else:
            # .pyスクリプトとして実行されている場合のパス設定
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
            print(
                "いくつかのファイルでエラーが発生しました。詳細は上記のログを確認してください。"
            )

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
                    with open(output_filepath, "w", encoding="utf-8") as f:
                        f.write(
                            component.CodeModule.Lines(
                                1, component.CodeModule.CountOfLines
                            )
                        )

            workbook.Close(SaveChanges=False)
            print(f"  [完了] {os.path.basename(excel_filepath)}")
            return True
        except Exception as e:
            print(
                f"  [エラー] {os.path.basename(excel_filepath)} の処理中にエラーが発生: {e}"
            )
            return False
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception as e:
                    # Quitに失敗してもプログラムは止めない
                    print(
                        f"  [警告] Excelの終了処理中にエラーが発生しました: {e}",
                        file=sys.stderr,
                    )

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
