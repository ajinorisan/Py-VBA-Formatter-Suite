# VBA Exporter

## 概要 (Overview)

<details>
<summary><strong>日本語</strong></summary>

---

**VBA Exporter** は、Excelファイル内に埋め込まれたVBAプロジェクトを、Gitなどのバージョン管理システムで扱いやすいように、個別のテキストファイル (`.bas`, `.cls`, `.frm`)として一括でエクスポートするツールです。

さらに、エクスポートと同時に**内蔵された高機能フォーマッターがVBAコードを自動で美しく整形**します。これにより、コードの可読性が向上し、チームでのコードレビューや差分（Diff）の確認が格段に容易になります。

### 主な機能

-   **VBAコードのエクスポート**: 標準モジュール、クラスモジュール、フォームモジュールを、元のコンポーネント名を維持したままファイルに出力します。
-   **自動コードフォーマット**: エクスポートと同時に、ネストされた複雑なブロック構造も含むVBAコードのインデントを正確に整形します。
-   **直感的なGUI操作**: 使いやすいGUIウィンドウから、ファイル選択ダイアログを開いて操作できます。
-   **複数ファイルの一括処理**: 複数のExcelファイルを一度に選択し、まとめてエクスポート処理を実行できます。
-   **リアルタイムログ表示**: 処理の進捗や結果がGUIウィンドウにリアルタイムで表示されます。

### スクリーンショット

![vba_exporter_jp](https://github.com/user-attachments/assets/4c42680b-9fff-4cff-84ee-af3fbd099cc9)

### 動作環境

-   **OS**: Windows 10 / 11 (本ツールはWindows専用です)
-   **アプリケーション**: Microsoft Excel

### 使い方

#### 実行ファイル (.exe) を使う

Pythonの環境構築が不要なため、ほとんどのユーザーにこの方法を推奨します。

1.  **ファイルのダウンロード**
    -   本リポジトリの [**Releasesページ**](https://github.com/TC-AJINORI/Py-VBA-Formatter-Suite/tree/main/_Releases) にアクセスします。
    -   最新バージョンのアセットから `vba_exporter.exe` をダウンロードします。

2.  **セキュリティに関する重要な注意**
    -   本プログラムは開発者によるデジタル署名が行われていません。そのため、ダウンロード時や実行時に **Windows Defender SmartScreen** やお使いのアンチウイルスソフトによって警告が表示される場合があります。
    -   これは、未知の実行ファイルに対する標準的な保護機能であり、必ずしもウイルスを意味するものではありません。
    -   実行するには、以下のように操作してください。
        -   Windowsの警告画面で **詳細情報** をクリックします。
        -   次に表示される **実行** ボタンをクリックします。
    -   本プログラムのダウンロードおよび実行は、これらのリスクを理解した上で、**自己の責任において**行ってください。

3.  **ツールの起動**
    -   ダウンロードした `vba_exporter.exe` をダブルクリックして実行します。
    -   表示されたウィンドウのボタンをクリックし、エクスポートしたいExcelファイルを選択します。

### 注意事項

-   エクスポートされたVBAコードは、実行元のフォルダ配下に `vba_source` という名前のフォルダが作成され、その中に保存されます。
-   VBAプロジェクトがパスワードで保護されている場合、コードの読み書きがブロックされるため、本ツールは機能しません。

### ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は `LICENSE` ファイルをご覧ください。

---
</details>

<br>

<details>
<summary><strong>English</strong></summary>

---

**VBA Exporter** is a tool that batch exports the VBA project embedded within an Excel file into individual text files (`.bas`, `.cls`, `.frm`) for easy handling in version control systems like Git.

Furthermore, upon export, the **built-in advanced formatter automatically beautifies the VBA code**. This improves code readability and makes code reviews and diff checking in a team environment significantly easier.

### Features

-   **VBA Code Export**: Exports standard modules, class modules, and form modules to files, maintaining their original component names.
-   **Automatic Code Formatting**: Simultaneously formats the indentation of VBA code, including complex nested block structures, upon export.
-   **Intuitive GUI Operation**: Allows users to operate via a user-friendly GUI window, opening a file selection dialog.
-   **Batch Processing of Multiple Files**: Supports selecting and processing multiple Excel files at once.
-   **Real-time Log Display**: Shows the progress and results of the processing in real-time in the GUI window.

### Screenshot

![vba_exporter_en](https://github.com/user-attachments/assets/4c42680b-9fff-4cff-84ee-af3fbd099cc9)

### System Requirements

-   **OS**: Windows 10 / 11 (This tool is for Windows only)
-   **Application**: Microsoft Excel

### Usage

#### Using the executable file (.exe)

This method is recommended for most users as it does not require setting up a Python environment.

1.  **Download the File**
    -   Access the [**Releases page**](https://github.com/TC-AJINORI/Py-VBA-Formatter-Suite/tree/main/_Releases) of this repository.
    -   Download `vba_exporter.exe` from the assets of the latest version.

2.  **Important Security Note**
    -   This program is not digitally signed by the developer. Therefore, you may see warnings from **Windows Defender SmartScreen** or your antivirus software when downloading or running it.
    -   This is a standard protection feature for unknown executables and does not necessarily mean it is a virus.
    -   To run it, please follow these steps:
        -   On the Windows warning screen, click **More info**.
        -   Then, click the **Run anyway** button that appears.
    -   Please download and run this program **at your own risk**, understanding these factors.

3.  **Launch the Tool**
    -   Double-click the downloaded `vba_exporter.exe` to run it.
    -   Click the button in the displayed window to select the Excel files you want to export.

### Notes

-   The exported VBA code is saved in a folder named `vba_source` created under the directory where the tool was executed.
-   If a VBA project is password-protected, this tool will not function as code reading and writing will be blocked.

### License

This project is licensed under the MIT License - see the `LICENSE` file for details.

---
</details>
