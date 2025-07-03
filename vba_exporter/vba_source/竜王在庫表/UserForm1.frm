Option Explicit

Private Sub CommandButton1_Click()    '商品CD列検索

    Dim aws_name As String
    aws_name = ActiveSheet.Name

    If aws_name = "在庫表" Then
        Call search_product_code(aws_name, "D", UserForm1)

    ElseIf aws_name = "予約" Then
        Call search_product_code(aws_name, "B", UserForm1)

    End If

    Unload UserForm1

End Sub

Private Sub CommandButton2_Click()

    Dim aws_name As String
    aws_name = ActiveSheet.Name

    If aws_name = "在庫表" Then
        Call filter_product_code(aws_name, "D", UserForm1)

    ElseIf aws_name = "予約" Then
        Call filter_product_code(aws_name, "B", UserForm1)

    End If

    Unload UserForm1

End Sub

Private Sub CommandButton3_Click()    '在庫シートJAN検索

    Dim aws_name As String
    aws_name = ActiveSheet.Name

    If aws_name = "在庫表" Then
        Call search_product_jancode(aws_name, "C", UserForm1)
        
    ElseIf aws_name = "予約" Then
        Call search_product_jancode(aws_name, "A", UserForm1)

    End If

    Unload UserForm1

End Sub


Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'入力制限をする。
    Select Case KeyAscii.Value
    Case 32 To 47    '記号
    Case 48 To 57    '0～9
    Case 65 To 90    'A～Z
    Case 97 To 122    'a～z
        KeyAscii.Value = KeyAscii.Value - 32    '大文字に強制変換
    Case Else
        KeyAscii.Value = 0
    End Select
End Sub

'Private Sub CommandButton1_Click()    '商品CD列検索
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "manual")
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    Dim text As String
'    text = UCase(UserForm1.TextBox1.text)
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'
'    With ws1
'
'        '
'        .AutoFilterMode = False
'        .Range(.Cells(4, 1), .Cells(4, ws1_4_row_lastcol)).AutoFilter
'
'        If ActiveCell.Column <> 4 Or Selection.CountLarge <> 1 Then
'            .Range("D4").Select
'        End If
'
'        On Error Resume Next
'        .Range("D4:D" & ws1_lastrow_d_col).Find(what:=text, _
          '                                                after:=ActiveCell, LookIn:=xlFormulas, LookAt _
          '                                                                                       :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
          '                                                False, MatchByte:=False, SearchFormat:=False).Activate
'
'        On Error GoTo 0
'
'
'        Dim syukka_day As Long
'        syukka_day = WorksheetFunction.Match(ws1.Range("D3"), ws1.Rows("4:4"), 0)
'
'        With ActiveWindow
'            .ScrollRow = ActiveCell.Row - 3
'            .ScrollColumn = syukka_day
'        End With
'
'        ws1.Cells(ActiveCell.Row, syukka_day).Select
'
'        Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'        Call ThisWorkbook.toggle_screen_update(True)
'
'        Call ThisWorkbook.set_calculation_mode(ws1, "automatic")
'
'        Unload UserForm1
'
'
'
'    End With
'
'End Sub

'Private Sub CommandButton2_Click()
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "manual")
'
'    Dim text As String
'    text = UCase(UserForm1.TextBox1.text)
'
'    Dim syouhincd_col As Long
'    syouhincd_col = 4
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "manual")
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    With ws1
'        .AutoFilterMode = False
'        .Range(.Cells(4, 1), .Cells(4, ws1_4_row_lastcol)).AutoFilter Field:=syouhincd_col, _
          '                                                                      Criteria1:="*" & text & "*", Operator:=xlFilterValues
'    End With
'
'
'    Dim syukka_day As Long
'    syukka_day = WorksheetFunction.Match(ws1.Range("D3"), ws1.Rows("4:4"), 0)
'
'    With ActiveWindow
'        .ScrollRow = ActiveCell.Row - 3
'        .ScrollColumn = syukka_day
'    End With
'
'    ws1.Cells(ActiveCell.Row, syukka_day).Select
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'    Call ThisWorkbook.toggle_screen_update(True)
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "automatic")
'
'    Unload UserForm1
'
'
'End Sub

'Private Sub CommandButton3_Click()    '在庫シートJAN検索
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "manual")
'
'    Dim text As Variant
'    text = UserForm1.TextBox1.text
'
'    Dim jancd_col As Long
'    jancd_col = 3
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "manual")
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    With ws1
'        .AutoFilterMode = False
'        .Range(.Cells(4, 1), .Cells(4, ws1_4_row_lastcol)).AutoFilter
'        If ActiveCell.Column <> 3 Or Selection.CountLarge <> 1 Then
'            .Range("C4").Select
'        End If
'    End With
'
'    Dim syukka_day As Long
'    syukka_day = WorksheetFunction.Match(ws1.Range("D3"), ws1.Rows("4:4"), 0)
'
'    Dim i As Long
'    Dim cell_value As String
'
'    Dim data As Variant
'    data = ws1.Range("C4:C" & ws1_lastrow_d_col).Value
'
'    For i = 1 To UBound(data, 1)
'        cell_value = CStr(data(i, 1))
'        If InStr(cell_value, text) > 0 Then
'            ws1.Cells(i + 3, syukka_day).Select
'
'            With ActiveWindow
'                .ScrollRow = ActiveCell.Row - 3
'                .ScrollColumn = syukka_day
'            End With
'
'            Exit For
'        End If
'    Next i
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'    Call ThisWorkbook.toggle_screen_update(True)
'
'    Call ThisWorkbook.set_calculation_mode(ws1, "automatic")
'
'    Unload UserForm1
'
'End Sub