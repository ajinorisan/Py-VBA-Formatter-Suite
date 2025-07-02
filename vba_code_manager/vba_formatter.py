import re

# --- ▼ フォーマッターのコアロジック ▼ ---

INDENT_KEYWORDS = [
    "If", "For", "Do", "With", "Sub", "Function", "Property", "Select", "Type"
]
DEDENT_KEYWORDS = [
    "End If", "Next", "Loop", "End With", "End Sub", "End Function", "End Property", "End Select", "End Type"
]
OUTDENT_KEYWORDS = [
    "Else", "ElseIf", "Case"
]

def format_vba_indent(code_string: str, indent_char: str = "    ") -> str:
    """
    VBAコードの文字列を受け取り、インデントを整形して返す。
    :param code_string: 元のVBAコード文字列
    :param indent_char: インデントに使う文字（例: スペース4つ）
    :return: 整形後のVBAコード文字列
    """
    lines = code_string.splitlines()
    formatted_lines = []
    current_indent_level = 0

    for line in lines:
        # 行頭と行末の空白を削除
        stripped_line = line.strip()

        # 空行はそのまま追加
        if not stripped_line:
            formatted_lines.append("")
            continue

        # コメント行は現在のインデントでそのまま追加
        if stripped_line.startswith("'"):
            formatted_lines.append(indent_char * current_indent_level + stripped_line)
            continue
        
        # 1行にまとまった "If ... Then ... End If" などを考慮しないシンプルな判定
        # 行の最初の単語を取得（大文字小文字を区別しないように、先頭を大文字化して比較）
        # 正規表現で行頭の単語を安全に取得
        match = re.match(r"^\s*([a-zA-Z_][a-zA-Z0-9_]*)", stripped_line)
        first_word = ""
        if match:
            # VBEの挙動に合わせて先頭を大文字化
            first_word_raw = match.group(1)
            first_word = first_word_raw.capitalize()
            # "End If" や "Select Case" のような複合キーワードに対応
            if first_word_raw.lower() == "end" and len(stripped_line.split()) > 1:
                first_word = " ".join(stripped_line.split()[:2]).capitalize()
            if first_word_raw.lower() == "select" and len(stripped_line.split()) > 1:
                if stripped_line.split()[1].lower() == "case":
                    first_word = "Select Case"


        # --- インデントレベルの調整 ---
        
        # Dedentキーワードかチェック
        if any(keyword.lower() == first_word.lower() for keyword in DEDENT_KEYWORDS):
            current_indent_level = max(0, current_indent_level - 1)

        # Outdentキーワードかチェック
        is_outdent = any(keyword.lower() == first_word.lower() for keyword in OUTDENT_KEYWORDS)
        if is_outdent:
            current_indent_level = max(0, current_indent_level - 1)
            
        # --- 行の出力 ---
        
        formatted_lines.append(indent_char * current_indent_level + stripped_line)
        
        # --- インデントレベルの事後調整 ---
        
        # Outdentキーワードだったらインデントを元に戻す
        if is_outdent:
            current_indent_level += 1

        # Indentキーワードかチェック
        if any(keyword.lower() == first_word.lower() for keyword in INDENT_KEYWORDS):
            # ただし、一行で完結するIf文はインデントしない
            # "If ... Then ..." の後にコードが続く場合はインデントしない
            if not (first_word.lower() == "if" and "then" in stripped_line.lower() and len(stripped_line) > stripped_line.lower().find("then") + 4):
                 current_indent_level += 1

    return "\n".join(formatted_lines)

# --- ▼ テスト用のコード ▼ ---
if __name__ == "__main__":
    test_code = """
Sub MyTest()
Dim i As Long
For i = 1 To 10
If i Mod 2 = 0 Then
Debug.Print "Even: " & i
Else
Debug.Print "Odd: " & i
End If
Next i
End Sub
"""
    formatted = format_vba_indent(test_code)
    print("--- Original ---")
    print(test_code)
    print("\n--- Formatted ---")
    print(formatted)