Option Explicit

Public Sub 数式コピー()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")

    Dim i As Long
    Dim col As Long

    Dim columns_copy As Variant
    columns_copy = Array("K", "L", "M", "O", "P", "Q", "R", "S", "T", "U", "X", "AA", "AB", "AC", "BX", "FF", "LJ", "LL")

    Dim ws1_lastrow_d_col As Long
    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "manual")
    Call toggle_sheet_protection(ws1, "un_protect")

    For i = LBound(columns_copy) To UBound(columns_copy)
        col = ws1.Range(columns_copy(i) & "1").Column
        ws1.Cells(5, col).Copy
        ws1.Range(ws1.Cells(6, col), ws1.Cells(ws1_lastrow_d_col - 1, col)).PasteSpecial Paste:=xlPasteFormulas

    Next i
    Application.CutCopyMode = False

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws1, "automatic")
    Call toggle_sheet_protection(ws1, "protect")


    Call toggle_screen_update(False)
    Call set_calculation_mode(ws2, "manual")
    Call toggle_sheet_protection(ws2, "un_protect")

    columns_copy = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")

    For i = LBound(columns_copy) To UBound(columns_copy)
        col = ws2.Range(columns_copy(i) & "1").Column
        ws2.Cells(5, col).Copy
        ws2.Range(ws2.Cells(6, col), ws2.Cells(ws1_lastrow_d_col - 1, col)).PasteSpecial Paste:=xlPasteFormulas

    Next i
    Application.CutCopyMode = False

    Dim search_strs As Variant
    search_strs = Array("合計", "入荷予定日", "生産予定", "次回予約合計", "次回入荷後引当", "入荷予定日2", "生産予定2", "次々回予約合計", "次々回入荷後引当")

    Dim ws2_4_row_lastcol As Long
    ws2_4_row_lastcol = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

    Dim match_index As Variant

    For i = LBound(search_strs) To UBound(search_strs)
        match_index = Application.Match(search_strs(i), ws2.Rows(4), 0)
        If Not IsError(match_index) Then
            col = match_index
            ws2.Cells(5, col).Copy
            ws2.Range(ws2.Cells(6, col), ws2.Cells(ws1_lastrow_d_col - 1, col)).PasteSpecial Paste:=xlPasteFormulas
        End If
    Next i

    Application.CutCopyMode = False

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws2, "automatic")
    Call toggle_sheet_protection(ws2, "protect")

End Sub



