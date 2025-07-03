# Py VBA Formatter Suite

**Py VBA Formatter Suite** は、Pythonで開発された、Excel VBAのコーディングとバージョン管理を近代化するためのツール群です。

VBA開発における「コード整形の手間」と「バージョン管理の難しさ」という2つの大きな課題を解決し、開発者がより創造的な作業に集中できる環境を提供します。

---

## ツール一覧 (Tools)

このスイートには、以下の2つのツールが含まれています。

### 1. Active VBA Formatter (リアルタイム・フォーマッター)

![Designer](https://github.com/user-attachments/assets/6c220dd4-2d5d-4601-a121-680cac8be2ed)

バックグラウンドで起動し、現在作業中のExcelファイルを常時監視します。VBEでコードを保存 (`Ctrl+S`) するたびに、**瞬時にコードのインデントを美しく整形**します。

手動でのインデント調整から解放され、思考を中断することなくコーディングを続けられます。

**[>> Active VBA Formatter の詳細はこちら (README)](./active_vba_formatter/README.md)**

### 2. VBA Exporter (VBA-Git連携ツール)

![VBA_Exporter](https://github.com/user-attachments/assets/e205aac2-6686-4076-abfa-ee8cadbe610e)

Excelファイル内のVBAプロジェクト（標準モジュール、クラス、フォーム）を、**個別のテキストファイルとして一括でエクスポート**します。

エクスポートされたファイルはGitなどのバージョン管理システムで差分を明確に追跡できるため、チームでの共同開発や変更履歴の管理が格段に容易になります。エクスポート時には自動でコード整形も行われます。

**[>> VBA Exporter の詳細はこちら (README)](./vba_exporter/README.md)**

---

## プロジェクトの目的 (Project Goal)

このプロジェクトは、VBAという強力なツールを、現代的な開発プラクティスと融合させることを目指しています。

-   **品質向上**: 整形されたコードは可読性が高く、バグの発見を容易にします。
-   **生産性向上**: 面倒な手作業を自動化し、開発者が本来の業務に集中できるようにします。
-   **共同作業の円滑化**: Gitを用いたバージョン管理を可能にし、チーム開発の基盤を整えます。

## ライセンス (License)

このプロジェクトはMITライセンスの下で公開されています。詳細は `LICENSE` ファイルをご覧ください。

This project is licensed under the MIT License - see the `LICENSE` file for details.
