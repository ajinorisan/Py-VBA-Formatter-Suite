Option Explicit

'Sub 商品別売上集計()
'
'Dim WS1 As Worksheet
'Set WS1 = ThisWorkbook.Worksheets("貼付")
'
'Dim WS2 As Worksheet
'Set WS2 = ThisWorkbook.Worksheets("部門別")
'
'Dim WS3 As Worksheet
'Set WS3 = ThisWorkbook.Worksheets("部門別商品別")
'
'Dim WS4 As Worksheet
'Set WS4 = ThisWorkbook.Worksheets("商品別")
'
'Dim WS5 As Worksheet
'Set WS5 = ThisWorkbook.Worksheets("前期")
'
'Dim WS6 As Worksheet
'Set WS6 = ThisWorkbook.Worksheets("商品分類マスタ")
'
'With ThisWorkbook.Worksheets.Add()
'        .Name = "TEMP"
'    End With
'
'Dim WS100 As Worksheet
'Set WS100 = ThisWorkbook.Worksheets("TEMP")
'
'Dim AWS1 As Worksheet
'
'Dim FOR1 As Long
'
'Dim 売上明細 As String
'売上明細 = "C:\Users\toyocase\Desktop\日報_AO\DATA\売上明細.xlsx"
'
'Dim 日報_AO As String
'日報_AO = "C:\Users\toyocase\Desktop\日報_AO\"
'
'Dim WS1_AC_LR As Long
'
'Dim WS100_AC_LR  As Long
'
'Dim WS6_AC_LR As Long
'
'Dim WS6_CC_LR As Long
'
'Dim WS5_AC_LR As Long
'
'If Dir(売上明細) = "" Then
'
'        MsgBox 売上明細 & vbCrLf & _
'               "が存在しません"
'
'               Exit Sub
'    Else
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlCalculationManual
'
'    WS1.Cells.ClearContents
'
'Workbooks.Open Filename:= _
'        売上明細, ReadOnly:=True
'
'        End If
'
'Set AWS1 = Workbooks("売上明細.xlsx").Worksheets("sheet1")
'
'AWS1.Cells.Copy
'WS1.Cells.PasteSpecial
'
'Application.CutCopyMode = False
'Workbooks("売上明細.xlsx").Close savechanges:=False
'
'WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
'
'WS1.Cells(1, 47).Formula = "売上月"
'
'With WorksheetFunction
'For FOR1 = 2 To WS1_AC_LR
'WS1.Cells(FOR1, 7).Formula = .Clean(WS1.Cells(FOR1, 7))
'WS1.Cells(FOR1, 47).Value = Mid(WS1.Cells(FOR1, 1), 5, 2) * 1
'Next FOR1
'End With
'
'WS1.Range("G:G").Copy
'WS100.Range("A:A").PasteSpecial
'
'Application.CutCopyMode = False
'
'    WS100.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
'
''    Dim WS100_KARA3 As Long
''
''    With WorksheetFunction
''    WS100_KARA3 = .Match("~03", WS100.Range("A:A"), 0)
''    End With
''    WS100.Rows(WS100_KARA3 & ":" & WS100_KARA3).Delete
'
'    WS100.Sort.SortFields.Clear
'    WS100.Sort.SortFields.Add Key:=WS100.Range("A:A") _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With WS100.Sort
'        .SetRange WS100.Range("A:A")
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    WS6_AC_LR = WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Row
'
' WS100_AC_LR = WS100.Cells(WS100.Rows.Count, 1).End(xlUp).Row
'
' With WorksheetFunction
'For FOR1 = 2 To WS100_AC_LR
'If .CountIf(WS6.Range("A:A"), WS100.Cells(FOR1, 1)) = 0 Then
'WS6.Cells(WS6_AC_LR + 1, 1).Value = WS100.Cells(FOR1, 1)
'On Error Resume Next
'WS6.Cells(WS6_AC_LR + 1, 2).Value = .VLookup(WS6.Cells(WS6_AC_LR + 1, 1), WS1.Range("G:H"), 2, False)
'On Error GoTo 0
'WS6_AC_LR = WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Row
'End If
'Next FOR1
'End With
'
'WS6_AC_LR = WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Row
'WS6_CC_LR = WS6.Cells(WS6.Rows.Count, 3).End(xlUp).Row
'WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row
'
'If WS6_AC_LR = WS6_CC_LR Then
'Application.DisplayAlerts = False
'WS100.Delete
'Application.DisplayAlerts = True
'
'With WorksheetFunction
'
'For FOR1 = 2 To WS1_AC_LR
'On Error Resume Next
'WS1.Cells(FOR1, 48).Value = .VLookup(WS1.Cells(FOR1, 7), WS6.Range("A:C"), 3, False)
'On Error GoTo 0
'Next FOR1
'
'End With
'
'For FOR1 = 2 To WS1_AC_LR
'If WS1.Cells(FOR1, 48) = "" Then
''On Error Resume Next
'WS1.Cells(FOR1, 48).Formula = "その他"
''On Error GoTo 0
'End If
'Next FOR1
'
'WS2.Calculate
'WS3.Calculate
'WS4.Calculate
'
'   With WorksheetFunction
'   If .Sum(WS1.Range("O:O")) <> WS3.Range("R96") Then
'
'   MsgBox "合計が合いません"
'
'   Application.ScreenUpdating = True
'   Exit Sub
'
'   Else
'  Workbooks.Add
'
'   Dim AWB As Workbook
'   Set AWB = ActiveWorkbook
''   AWB.Name = Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_部門売上集計表"
'
'   ThisWorkbook.Sheets(Array("部門別", "部門別商品別", "商品別")).Copy Before:=AWB.Sheets(1)
'
'   Dim AWBS1 As Worksheet
'   Set AWBS1 = AWB.Worksheets("部門別")
'   Dim AWBS2 As Worksheet
'   Set AWBS2 = AWB.Worksheets("部門別商品別")
'   Dim AWBS3 As Worksheet
'   Set AWBS3 = AWB.Worksheets("商品別")
'
'   AWBS3.Cells.Copy
'   AWBS3.Cells.PasteSpecial Paste:=xlPasteValues
' AWBS2.Cells.Copy
'   AWBS2.Cells.PasteSpecial Paste:=xlPasteValues
'   AWBS1.Cells.Copy
'   AWBS1.Cells.PasteSpecial Paste:=xlPasteValues
'
'Application.DisplayAlerts = False
'
'  ActiveWorkbook.SaveAs Filename:= _
' 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_部門売上集計表"
'
'AWB.Close savechanges:=False
'Application.DisplayAlerts = True
'End If
'End With
'
'Else
'Application.DisplayAlerts = False
'WS100.Delete
'Application.DisplayAlerts = True
'Application.ScreenUpdating = True
'WS6.Activate
'WS6.Cells(WS6_CC_LR + 1, 3).Activate
'MsgBox "部門分類入力"
'Exit Sub
'End If
'
'
'
'Application.Calculation = xlCalculationAutomatic
'MsgBox "更新完了"
'WS2.Activate
'Application.ScreenUpdating = True
'End Sub