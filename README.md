# Py VBA Formatter Suite

![Tool Icon](https://github.com/user-attachments/asset47a2b437-be64-4e21-a10f-6a7d7e98b405)
A tool suite to modernize Excel VBA development, including a real-time code formatter and a Git-friendly exporter.

<br>

<details>
<summary><strong>🇯🇵 日本語 (Japanese)</strong></summary>

---

**Py VBA Formatter Suite** は、Pythonで開発された、Excel VBAのコーディングとバージョン管理を近代化するためのツール群です。

VBA開発における「コード整形の手間」と「バージョン管理の難しさ」という2つの大きな課題を解決し、開発者がより創造的な作業に集中できる環境を提供します。

### ツール一覧

このスイートには、以下の2つのツールが含まれています。

#### 1. Active VBA Formatter (リアルタイム・フォーマッター)
バックグラウンドで起動し、現在作業中のExcelファイルを常時監視します。VBEでコードを保存 (`Ctrl+S`) するたびに、**瞬時にコードのインデントを美しく整形**します。手動でのインデント調整から解放され、思考を中断することなくコーディングを続けられます。

**[>> Active VBA Formatter の詳細はこちら](./active_vba_formatter/README.md)**

#### 2. VBA Exporter (VBA-Git連携ツール)
![VBA Exporter](https://github.com/user-attachments/asseta0293b79-7e86-4b7c-9ab1-4a391d822cee)

Excelファイル内のVBAプロジェクト（標準モジュール、クラス、フォーム）を、**個別のテキストファイルとして一括でエクスポート**します。エクスポートされたファイルはGitなどのバージョン管理システムで差分を明確に追跡できるため、チームでの共同開発や変更履歴の管理が格段に容易になります。エクスポート時には自動でコード整形も行われます。

**[>> VBA Exporter の詳細はこちら](./vba_exporter/README.md)**

### プロジェクトの目的

このプロジェクトは、VBAという強力なツールを、現代的な開発プラクティスと融合させることを目指しています。

-   **品質向上**: 整形されたコードは可読性が高く、バグの発見を容易にします。
-   **生産性向上**: 面倒な手作業を自動化し、開発者が本来の業務に集中できるようにします。
-   **共同作業の円滑化**: Gitを用いたバージョン管理を可能にし、チーム開発の基盤を整えます。

---

</details>

<br>

<details>
<summary><strong>🇬🇧 English (英語)</strong></summary>

---

**Py VBA Formatter Suite** is a collection of tools developed in Python to modernize Excel VBA coding and version control.

It solves two major challenges in VBA development—the hassle of code formatting and the difficulty of version control—providing an environment where developers can focus on more creative tasks.

### Tools Overview

This suite includes the following two tools:

#### 1. Active VBA Formatter (Real-time Formatter)
Runs in the background and constantly monitors the Excel file you are currently working on. Every time you save your code in the VBE (`Ctrl+S`), it **instantly formats the code indentation beautifully**. This frees you from manual indentation adjustments, allowing you to code without interrupting your train of thought.

**[>> Click here for Active VBA Formatter details](./active_vba_formatter/README.md)**

#### 2. VBA Exporter (VBA-Git Integration Tool)
![VBA Exporter](https://github.com/user-attachments/asseta0293b79-7e86-4b7c-9ab1-4a391d822cee)

Exports the entire VBA project (standard modules, classes, forms) from an Excel file into **individual text files**. These exported files can be clearly tracked for differences in version control systems like Git, making team collaboration and change history management significantly easier. Code formatting is also performed automatically during export.

**[>> Click here for VBA Exporter details](./vba_exporter/README.md)**

### Project Goal

This project aims to merge the powerful tool of VBA with modern development practices.

-   **Improved Quality**: Well-formatted code is highly readable and makes bug detection easier.
-   **Increased Productivity**: Automates tedious manual tasks, allowing developers to concentrate on their primary work.
-   **Smoother Collaboration**: Enables version control using Git, laying the groundwork for team development.

---

</details>

<br>

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.