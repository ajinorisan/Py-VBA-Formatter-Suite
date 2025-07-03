Option Explicit

Public Sub 簡易在庫作成()


    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")
    Set ws3 = ThisWorkbook.Worksheets("受注残")

    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "manual")
    Call toggle_sheet_protection(ws1, "un_protect")
    Call set_calculation_mode(ws2, "manual")
    Call toggle_sheet_protection(ws2, "un_protect")
    
    Application.DisplayAlerts = False

    Dim kani_path As String

    kani_path = _
                Environ("USERPROFILE") & "\Desktop\" & Format(Date, "yymmdd") & "簡易在庫表.xlsx"

    Workbooks.Add
    Dim awb As Workbook
    Set awb = ActiveWorkbook
    awb.SaveAs Filename:=kani_path

    awb.Worksheets("sheet1").Name = "在庫数"
    Dim aws1 As Worksheet
    Set aws1 = awb.Worksheets("在庫数")


    Dim ws1_4_row_last_col As Long
    ws1_4_row_last_col = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

    If ws1.AutoFilterMode = True Then
        ws1.AutoFilterMode = False
        ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_4_row_last_col)).AutoFilter
    End If

    With ws1
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    ws1.Range("A:AC").Copy
    aws1.Range("A:AC").PasteSpecial Paste:=xlPasteFormats
    aws1.Range("A:AC").PasteSpecial Paste:=xlPasteValues
    aws1.Cells.FormatConditions.Delete


    Dim ws2_4_row_last_col As Long
    ws2_4_row_last_col = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

    If ws2.AutoFilterMode = True Then
        ws2.AutoFilterMode = False
        ws2.Range(ws2.Cells(4, 1), ws2.Cells(4, ws2_4_row_last_col)).AutoFilter
    End If

    With ws2
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    ws2.Range(ws2.Cells(1, 13), ws2.Cells(1, ws2_4_row_last_col)).EntireColumn.Copy
    aws1.Range("AD1").PasteSpecial Paste:=xlPasteFormats
    aws1.Range("AD1").PasteSpecial Paste:=xlPasteValues
    aws1.Cells.FormatConditions.Delete

    aws1.Rows("2:2").RowHeight = 22.5
    aws1.Rows("3:3").RowHeight = 37.5
    aws1.Rows("4:4").RowHeight = 37.5
    
    Dim rng As Range

    Set rng = Union(aws1.Range("A2"), aws1.Range("D2"), aws1.Range("E2"), aws1.Range("D3"), aws1.Range("L3"), aws1.Range("A1:AC1"))
    rng.ClearContents
    
    aws1.Range("F5").Select

    ActiveWindow.FreezePanes = True
    
    awb.Save
awb.Close
    
    Application.DisplayAlerts = True

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws1, "automatic")
    Call toggle_sheet_protection(ws1, "protect")
    Call set_calculation_mode(ws2, "automatic")
    Call toggle_sheet_protection(ws2, "protect")
End Sub



'Sub 簡易在庫作成()
'
'    Application.EnableEvents = False
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'
'    Dim 竜王簡易在庫表フルパス As String
'    竜王簡易在庫表フルパス = "C:\Users\TOYOC-303\Desktop\" & Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日竜王簡易在庫表.xlsx"
'
'    Dim WS1 As Worksheet
'    Dim WS2 As Worksheet
'    Dim WS3 As Worksheet
'    Dim WS8 As Worksheet
'
'    Set WS1 = ThisWorkbook.Worksheets("在庫数")
'    Set WS2 = ThisWorkbook.Worksheets("予約")
'    Set WS3 = ThisWorkbook.Worksheets("受注残")
'    Set WS8 = ThisWorkbook.Worksheets("OEM在庫")
'
'    Workbooks.Add
'    Dim NB As Workbook
'    Set NB = ActiveWorkbook
'    NB.SaveAs fileName:=竜王簡易在庫表フルパス
'
'    Dim FOR1 As Long
'    Select Case NB.Worksheets.Count
'    Case 1
'        For FOR1 = 2 To 4
'            NB.Worksheets.Add after:=NB.Worksheets(NB.Worksheets.Count), Count:=1
'        Next FOR1
'    Case 2
'        For FOR1 = 3 To 4
'            NB.Worksheets.Add after:=NB.Worksheets(NB.Worksheets.Count), Count:=1
'        Next FOR1
'    Case 3
'        NB.Worksheets.Add after:=NB.Worksheets(NB.Worksheets.Count), Count:=1
'    Case Else
'    End Select
'
'    NB.Worksheets("sheet1").Name = "在庫数"
'    NB.Worksheets("sheet2").Name = "予約"
'    NB.Worksheets("sheet3").Name = "受注残"
'    NB.Worksheets("sheet4").Name = "OEM在庫"
'
'    Dim NB_WS1 As Worksheet
'    Dim NB_WS2 As Worksheet
'    Dim NB_WS3 As Worksheet
'    Dim NB_WS4 As Worksheet
'
'    Set NB_WS1 = NB.Worksheets("在庫数")
'    Set NB_WS2 = NB.Worksheets("予約")
'    Set NB_WS3 = NB.Worksheets("受注残")
'    Set NB_WS4 = NB.Worksheets("OEM在庫")
'
'    WS1.Unprotect
'
'    WS1.Columns.Hidden = False
'    WS1.Rows.Hidden = False
'    WS1.AutoFilterMode = False
'    WS1.Range("4:4").AutoFilter
'
'    WS1.Range("A:Y").Copy
'    NB_WS1.Range("A:Y").PasteSpecial Paste:=xlPasteFormats
'    NB_WS1.Range("A:Y").PasteSpecial Paste:=xlPasteValues
'    NB_WS1.Cells.FormatConditions.Delete
'    NB_WS1.Range("E2").ClearContents
'    NB_WS1.Activate
'    NB_WS1.Range("F5").Select
'    ActiveWindow.FreezePanes = True
'    NB_WS1.Range("1:1").Delete
'
'    WS1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '              , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '                AllowDeletingColumns:=True, AllowFiltering:=True
'
'    WS8.Unprotect
'
'    WS8.Columns.Hidden = False
'    WS8.Rows.Hidden = False
'    WS8.AutoFilterMode = False
'    WS8.Range("4:4").AutoFilter
'
'    WS8.Range("A:Y").Copy
'    NB_WS4.Range("A:Y").PasteSpecial Paste:=xlPasteFormats
'    NB_WS4.Range("A:Y").PasteSpecial Paste:=xlPasteValues
'    NB_WS4.Cells.FormatConditions.Delete
'    NB_WS4.Range("E2").ClearContents
'    NB_WS4.Activate
'    NB_WS4.Range("F5").Select
'    ActiveWindow.FreezePanes = True
'    NB_WS4.Range("1:1").Delete
'
'    WS8.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '              , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '                AllowDeletingColumns:=True, AllowFiltering:=True
'
'    WS2.Unprotect
'
'    WS2.Columns.Hidden = False
'    WS2.Rows.Hidden = False
'    WS2.AutoFilterMode = False
'    WS2.Range("4:4").AutoFilter
'
'    WS2.Cells.Copy
'    NB_WS2.Cells.PasteSpecial Paste:=xlPasteFormats
'    NB_WS2.Cells.PasteSpecial Paste:=xlPasteValues
'    NB_WS2.Activate
'    NB_WS2.Range("H5").Select
'    ActiveWindow.FreezePanes = True
'
'    WS2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '              , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '                AllowDeletingColumns:=True, AllowFiltering:=True
'
'    WS3.Unprotect
'
'    WS3.Cells.Copy
'    NB_WS3.Cells.PasteSpecial Paste:=xlPasteFormats
'    NB_WS3.Cells.PasteSpecial Paste:=xlPasteValues
'    NB_WS3.Activate
'    NB_WS3.Range("B4").Select
'    ActiveWindow.FreezePanes = True
'
'
'    NB.Save
'    NB.Close
'
'    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
'
'    Application.EnableEvents = True
'
'    MsgBox "デスクトップに簡易在庫表作成完了"
'
'End Sub
