# Active VBA Formatter

## 概要 (Overview)

<details>
<summary><strong>日本語</strong></summary>

---

**Active VBA Formatter** は、Excel VBAの開発効率を劇的に向上させるためのリアルタイム・コードフォーマッターです。

このツールはバックグラウンドで起動し、現在アクティブになっているExcelファイルを常時監視します。あなたがVBE（Visual Basic Editor）でコードを記述し、ファイルを保存するたびに、瞬時にコードのインデントを美しく整形します。

手動でのインデント調整という煩わしい作業から解放され、コーディングそのものに集中できる環境を提供します。

### 主な機能

-   **リアルタイム監視**: フォアグラウンドのExcelブックを自動で認識し、監視対象を動的に切り替えます。
-   **自動フォーマット**: VBAコードの保存 (`Ctrl+S`) を検知し、瞬時にインデントを整形します。
    -   `If`, `Select Case`, `For`, `Do`, `With`, `Sub`, `Function`, `Property` 等のブロック構造に対応。
    -   ネストされた複雑なブロック構造も正確に解析します。
-   **スマートな終了処理**: 監視対象のExcelがすべて終了すると、ツールを終了するか確認ダイアログを表示します。

### スクリーンショット

![active_vba_formatter_jp](https://github.com/user-attachments/assets/0031d017-c571-49c0-a427-37a4ae651631)

### 動作環境

-   **OS**: Windows 10 / 11 (本ツールはWindows専用です)
-   **アプリケーション**: Microsoft Excel

### 使い方

#### 実行ファイル (.exe) を使う (推奨)

Pythonの環境構築が不要なため、ほとんどのユーザーにこの方法を推奨します。

1.  **ファイルのダウンロード**
    -   本リポジトリの [**Releasesページ**](https://github.com/TC-AJINORI/Py-VBA-Formatter-Suite/tree/main/_Releases) にアクセスします。
    -   最新バージョンのアセットから `active_vba_formatter.exe` をダウンロードします。

2.  **セキュリティに関する重要な注意**
    -   本プログラムは開発者によるデジタル署名が行われていません。そのため、ダウンロード時や実行時に **Windows Defender SmartScreen** やお使いのアンチウイルスソフトによって警告が表示される場合があります。
    -   これは、未知の実行ファイルに対する標準的な保護機能であり、必ずしもウイルスを意味するものではありません。
    -   実行するには、以下のように操作してください。
        -   Windowsの警告画面で **詳細情報** をクリックします。
        -   次に表示される **実行** ボタンをクリックします。
    -   本プログラムのダウンロードおよび実行は、これらのリスクを理解した上で、**自己の責任において**行ってください。

3.  **ツールの起動**
    -   ダウンロードした `active_vba_formatter.exe` をダブルクリックして実行します。
    -   ツールがバックグラウンドで起動し、Excelの監視を開始します。

### 注意事項

-   VBAプロジェクトがパスワードで保護されている場合、コードの読み書きがブロックされるため、本ツールは機能しません。
-   Excelが「応答なし」の状態になると、COM接続エラーが発生し、監視が中断されることがあります。

### ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は `LICENSE` ファイルをご覧ください。

---
</details>

<br>

<details>
<summary><strong>English</strong></summary>

---

**Active VBA Formatter** is a real-time code formatter designed to dramatically improve the development efficiency of Excel VBA.

This tool runs in the background and constantly monitors the currently active Excel file. Every time you write code in the VBE (Visual Basic Editor) and save the file, it instantly and beautifully formats the code indentation.

It frees you from the tedious task of manual indentation, providing an environment where you can focus solely on coding.

### Features

-   **Real-time Monitoring**: Automatically recognizes the foreground Excel workbook and dynamically switches the monitoring target.
-   **Automatic Formatting**: Detects when VBA code is saved (`Ctrl+S`) and instantly formats the indentation.
    -   Supports block structures such as `If`, `Select Case`, `For`, `Do`, `With`, `Sub`, `Function`, and `Property`.
    -   Accurately parses complex nested block structures.
-   **Smart Exit Handling**: Displays a confirmation dialog to exit the tool when all monitored Excel windows are closed.

### Screenshot

![active_vba_formatter_en](https://github.com/user-attachments/assets/0031d017-c571-49c0-a427-37a4ae651631)

### System Requirements

-   **OS**: Windows 10 / 11 (This tool is for Windows only)
-   **Application**: Microsoft Excel

### Usage

#### Using the executable file (.exe) (Recommended)

This method is recommended for most users as it does not require setting up a Python environment.

1.  **Download the File**
    -   Access the [**Releases page**](https://github.com/TC-AJINORI/Py-VBA-Formatter-Suite/tree/main/_Releases) of this repository.
    -   Download `active_vba_formatter.exe` from the assets of the latest version.

2.  **Important Security Note**
    -   This program is not digitally signed by the developer. Therefore, you may see warnings from **Windows Defender SmartScreen** or your antivirus software when downloading or running it.
    -   This is a standard protection feature for unknown executables and does not necessarily mean it is a virus.
    -   To run it, please follow these steps:
        -   On the Windows warning screen, click **More info**.
        -   Then, click the **Run anyway** button that appears.
    -   Please download and run this program **at your own risk**, understanding these factors.

3.  **Launch the Tool**
    -   Double-click the downloaded `active_vba_formatter.exe` to run it.
    -   The tool will start in the background and begin monitoring Excel.

### Notes

-   If a VBA project is password-protected, this tool will not function as code reading and writing will be blocked.
-   If Excel becomes "Not Responding," a COM connection error may occur, and monitoring may be interrupted.

### License

This project is licensed under the MIT License - see the `LICENSE` file for details.

---
</details>
