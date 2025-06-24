Option Explicit

Public Sub toggle_screen_update(enable As Boolean)
    Application.ScreenUpdating = enable
End Sub

Sub toggle_sheet_protection(ws As Worksheet, action As String)
    If Not ws Is Nothing Then
        If action = "un_protect" Then

            If ws.ProtectContents Then
                ws.Unprotect
            End If
        ElseIf action = "protect" Then

            If Not ws.ProtectContents Then
                ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
                           AllowDeletingColumns:=True, AllowFiltering:=True
            End If
        End If
    End If
End Sub

Public Sub set_calculation_mode(ws As Worksheet, mode As String)

    If mode = "automatic" Then
        Application.Calculation = xlCalculationAutomatic
    ElseIf mode = "manual" Then
        Application.Calculation = xlCalculationManual
    End If

End Sub

Public Sub search_product_code(sheet_name As String, target_col As String, uf As UserForm)     ' 商品CD列検索

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheet_name)

    Dim ws_last_row_target_col As Long
    ws_last_row_target_col = ws.Cells(ws.Rows.Count, target_col).End(xlUp).Row

    Dim ws_4_row_last_col As Long
    ws_4_row_last_col = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    Call set_calculation_mode(ws, "manual")
    Call toggle_screen_update(False)

    Dim search_text As String
    search_text = UCase(uf.TextBox1.text)

    Call toggle_sheet_protection(ws, "un_protect")

    With ws
        .AutoFilterMode = False
        .Range(.Cells(4, 1), .Cells(4, ws_4_row_last_col)).AutoFilter

        If ActiveCell.Column <> 4 Or Selection.CountLarge <> 1 Then
            .Cells(4, target_col).Select
        End If

        Dim i As Long
        Dim cell_value As String
        Dim data As Variant
        data = ws.Range(.Cells(4, target_col), .Cells(ws_last_row_target_col, target_col)).Value

        For i = 1 To UBound(data, 1)    ' 1から始める
            cell_value = CStr(data(i, 1))
            If InStr(cell_value, search_text) > 0 Then


                If sheet_name = "在庫表" Then
                    Dim syukka_day As Long
                    syukka_day = WorksheetFunction.Match(CDbl(ws.Range("D3")), ws.Rows("4:4"), 0)
                    ws.Cells(i + 3, syukka_day).Select
                    With ActiveWindow
                        .ScrollRow = ActiveCell.Row - 3
                        .ScrollColumn = syukka_day
                    End With
                Else
                    ws.Cells(i + 3, target_col).Select
                    ActiveWindow.ScrollRow = ActiveCell.Row - 3
                End If

                Exit For
            End If
        Next i

        ActiveWindow.ScrollRow = ActiveCell.Row - 3

        Call toggle_sheet_protection(ws, "protect")
        Call toggle_screen_update(True)
        Call set_calculation_mode(ws, "automatic")

    End With
End Sub

Public Sub filter_product_code(sheet_name As String, target_col As String, uf As UserForm)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheet_name)

    Dim ws_last_row_target_col As Long
    ws_last_row_target_col = ws.Cells(ws.Rows.Count, target_col).End(xlUp).Row

    Dim ws_4_row_lastcol As Long
    ws_4_row_lastcol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    Dim search_text As String
    search_text = UCase(uf.TextBox1.text)

    Dim syouhincd_col As Long
    syouhincd_col = Columns(target_col).Column

    Call toggle_sheet_protection(ws, "un_protect")

    Call set_calculation_mode(ws, "manual")

    Call toggle_screen_update(False)

    With ws
        .AutoFilterMode = False
        .Range(.Cells(4, 1), .Cells(4, ws_4_row_lastcol)).AutoFilter Field:=syouhincd_col, _
                                                                     Criteria1:="*" & search_text & "*", Operator:=xlFilterValues
    End With



    If sheet_name = "在庫表" Then
        Dim syukka_day As Long
        syukka_day = WorksheetFunction.Match(CDbl(ws.Range("D3")), ws.Rows("4:4"), 0)
        ws.Cells(ActiveCell.Row, syukka_day).Select
        With ActiveWindow
            .ScrollRow = ActiveCell.Row - 3
            .ScrollColumn = syukka_day
        End With
    Else
        ws.Cells(ActiveCell.Row, syouhincd_col - 1).Select
        ActiveWindow.ScrollRow = ActiveCell.Row - 3
    End If


    Call toggle_sheet_protection(ws, "protect")

    Call toggle_screen_update(True)

    Call set_calculation_mode(ws, "automatic")

End Sub

Public Sub search_product_jancode(sheet_name As String, target_col As String, uf As UserForm)     'JANCD列検索

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheet_name)

    Dim ws_last_row_target_col As Long
    ws_last_row_target_col = ws.Cells(ws.Rows.Count, target_col).End(xlUp).Row

    Dim ws_4_row_last_col As Long
    ws_4_row_last_col = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    Call set_calculation_mode(ws, "manual")
    Call toggle_screen_update(False)

    Dim search_text As String
    search_text = UCase(uf.TextBox1.text)

    Call toggle_sheet_protection(ws, "un_protect")

    With ws
        .AutoFilterMode = False
        .Range(.Cells(4, 1), .Cells(4, ws_4_row_last_col)).AutoFilter

        If ActiveCell.Column <> 4 Or Selection.CountLarge <> 1 Then
            .Cells(4, target_col).Select
        End If


        Dim data As Variant
        data = ws.Range(.Cells(4, target_col), .Cells(ws_last_row_target_col, target_col)).Value

        Dim i As Long
        Dim cell_value As String

        For i = 1 To UBound(data, 1)
            cell_value = CStr(data(i, 1))
            If InStr(cell_value, search_text) > 0 Then

                If sheet_name = "在庫表" Then
                    Dim syukka_day As Long
                    syukka_day = WorksheetFunction.Match(ws.Range("D3"), ws.Rows("4:4"), 0)
                    ws.Cells(i + 3, syukka_day).Select
                    With ActiveWindow
                        .ScrollRow = ActiveCell.Row - 3
                        .ScrollColumn = syukka_day
                    End With
                Else
                    ws.Cells(i + 3, target_col).Select
                    ActiveWindow.ScrollRow = ActiveCell.Row - 3
                End If

                Exit For
            End If
        Next i

        Call toggle_sheet_protection(ws, "protect")
        Call toggle_screen_update(True)
        Call set_calculation_mode(ws, "automatic")

    End With

End Sub

Public Sub clean_vlue(ws As Worksheet, target_col As Long)
    Dim i As Long

    Dim ws_last_row_target_col As Long
    ws_last_row_target_col = ws.Cells(ws.Rows.Count, target_col).End(xlUp).Row

    Call set_calculation_mode(ws, "manual")
    Call toggle_screen_update(False)

    With WorksheetFunction
        For i = 2 To ws_last_row_target_col
            ws.Cells(i, target_col).Value = .Clean(ws.Cells(i, target_col))
        Next i
    End With

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws, "automatic")
End Sub