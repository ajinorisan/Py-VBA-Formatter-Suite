Option Explicit

'Sub 得意先別集計()
'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
'Dim TIME As Single
'
'TIME = Timer   '時間計測開始
'
'Dim WS5 As Worksheet
'Set WS5 = ThisWorkbook.Worksheets("前々期")
'
'Dim WS7 As Worksheet
'Set WS7 = ThisWorkbook.Worksheets("前期")
'
'
' With Worksheets.Add()
'      .Name = "TEMP"
'    End With
'
'  Dim WS2 As Worksheet
'  Set WS2 = ThisWorkbook.Worksheets("TEMP")
'
'Dim 売上明細 As String
'売上明細 = ThisWorkbook.Path & "\DATA\担当者・得意先別売上.csv"
'
'Workbooks.Open Filename:= _
'        売上明細, ReadOnly:=True
'
'        Dim WS3 As Worksheet
'        Set WS3 = ThisWorkbook.Worksheets("今期")
'
'        Dim AWS1 As Worksheet
'
'        Set AWS1 = Workbooks("担当者・得意先別売上.csv").Worksheets(1)
'
'        Dim AWS1_AC_LR As Long
'AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row
'
'        AWS1.Range("A:D").NumberFormatLocal = "@"
'
'        Dim FOR1 As Long
'With WorksheetFunction
'For FOR1 = 2 To AWS1_AC_LR - 1
'AWS1.Cells(FOR1, 1).Formula = .Text(AWS1.Cells(FOR1, 1), "000000")
'      AWS1.Cells(FOR1, 3).Formula = .Text(AWS1.Cells(FOR1, 3), "00000000")
'      AWS1.Cells(FOR1, 5).Value = Mid(AWS1.Cells(FOR1, 5), 5, 2)
'  Next FOR1
'  End With
'
'  AWS1.Range(AWS1_AC_LR & ":" & AWS1_AC_LR).Delete
'
'  AWS1.Cells.Copy
'  WS2.Cells.PasteSpecial
'
'  Application.DisplayAlerts = False
'  Workbooks("担当者・得意先別売上.csv").Close savechanges:=False
'  Application.DisplayAlerts = True
'
'  WS3.Cells.ClearContents
'
'Dim WS2_AC_LR As Long
'WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row
''
''  WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS3_AC_LR, 5)).Copy
''  WS2.Range(WS2.Cells(1, 1), WS2.Cells(WS3_AC_LR, 5)).PasteSpecial
'
'  With WorksheetFunction
'For FOR1 = 2 To WS2_AC_LR
'WS2.Cells(FOR1, 8).Value = _
'.SumIfs(WS2.Range("F1:F" & WS2_AC_LR), WS2.Range("C1:C" & WS2_AC_LR), WS2.Cells(FOR1, 3), WS2.Range("E1:E" & WS2_AC_LR), WS2.Cells(FOR1, 5))
'Next FOR1
'
'  WS2.Range(WS2.Cells(1, 1), WS2.Cells(WS2_AC_LR, 8)).RemoveDuplicates Columns:=Array(3, 5, 8), Header:=xlYes
'
'  WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row
'
'WS2.Range(WS2.Cells(1, 1), WS2.Cells(WS2_AC_LR, 5)).Copy
'WS3.Cells(1, 1).PasteSpecial
'WS2.Range(WS2.Cells(1, 8), WS2.Cells(WS2_AC_LR, 8)).Copy
'WS3.Cells(1, 6).PasteSpecial
'WS2.Range(WS2.Cells(1, 8), WS2.Cells(WS2_AC_LR, 8)).Copy
'WS2.Cells(1, 7).PasteSpecial
'
''WS2.Range("C:D").NumberFormatLocal = "@"
'
'Dim WS2_CC_LR As Long
'WS2_CC_LR = WS2.Cells(WS2.Rows.Count, 3).End(xlUp).Row
'
'Dim WS5_AC_LR As Long
'WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row
'
'For FOR1 = 2 To WS5_AC_LR
'WS5.Cells(FOR1, 17).Value = .CountIf(WS2.Range("C:C"), WS5.Cells(FOR1, 1))
'Next FOR1
'
'WS5.Range(WS5.Cells(1, 1), WS5.Cells(1, 17)).AutoFilter Field:=17, Criteria1:=0
'WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row
'
'WS5.Range(WS5.Cells(2, 1), WS5.Cells(WS5_AC_LR, 1)).Copy
'WS2.Cells(WS2_CC_LR + 1, 3).PasteSpecial Paste:=xlPasteValues
'
'WS2_CC_LR = WS2.Cells(WS2.Rows.Count, 3).End(xlUp).Row
'
'Dim WS7_AC_LR As Long
'WS7_AC_LR = WS7.Cells(WS7.Rows.Count, 1).End(xlUp).Row
'
'For FOR1 = 2 To WS7_AC_LR
'WS7.Cells(FOR1, 7).Value = .CountIf(WS2.Range("C:C"), WS7.Cells(FOR1, 3))
'Next FOR1
'
'For FOR1 = 2 To WS7_AC_LR
'WS7.Cells(FOR1, 7).Value = .CountIf(WS7.Range(WS7.Cells(2, 3), WS7.Cells(FOR1, 3)), WS7.Cells(FOR1, 3)) - 1
'Next FOR1
'
'On Error Resume Next
'WS7.Range(WS7.Cells(1, 1), WS7.Cells(1, 7)).AutoFilter Field:=7, Criteria1:=0
'WS7_AC_LR = WS7.Cells(WS7.Rows.Count, 1).End(xlUp).Row
'On Error GoTo 0
'
'WS7.Range(WS7.Cells(2, 3), WS7.Cells(WS7_AC_LR, 3)).Copy
'WS2.Cells(WS2_CC_LR + 1, 3).PasteSpecial Paste:=xlPasteValues
'
'WS2_CC_LR = WS2.Cells(WS2.Rows.Count, 3).End(xlUp).Row
'
'WS2.Range(WS2.Cells(WS2_AC_LR + 1, 6), WS2.Cells(WS2_CC_LR, 6)).Value = 0
'
'WS5.AutoFilterMode = False
'WS7.AutoFilterMode = False
'
'WS5.Range("Q:Q").Delete
'WS7.Range("G:G").Delete
'
'Dim 得意先マスタ As String
'得意先マスタ = ThisWorkbook.Path & "\DATA\得意先マスタ.xlsx"
'
'Workbooks.Open Filename:= _
'        得意先マスタ, ReadOnly:=True
'
'Set AWS1 = Workbooks("得意先マスタ.xlsx").Worksheets("sheet1")
'
'AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row
'
'        AWS1.Range("A:A").NumberFormatLocal = "@"
'        AWS1.Range("AM:AM").NumberFormatLocal = "@"
'
'       For FOR1 = 2 To AWS1_AC_LR
'
'        AWS1.Cells(FOR1, 1).Value = .Clean(AWS1.Cells(FOR1, 1))
'        AWS1.Cells(FOR1, 39).Value = .Clean(AWS1.Cells(FOR1, 39))
'        Next FOR1
'
'
'On Error Resume Next
'For FOR1 = 2 To WS2_CC_LR
'WS2.Cells(FOR1, 4).Value = .Index(AWS1.Cells, .Match(WS2.Cells(FOR1, 3), AWS1.Range("A:A"), 0), 4)
'WS2.Cells(FOR1, 1).Value = .Index(AWS1.Cells, .Match(WS2.Cells(FOR1, 3), AWS1.Range("A:A"), 0), 39)
'WS2.Cells(FOR1, 2).Value = .Index(AWS1.Cells, .Match(WS2.Cells(FOR1, 3), AWS1.Range("A:A"), 0), 40)
''WS2.Cells(FOR1, 7).Value = .Rank(WS2.Cells(FOR1, 6), WS2.Range(WS2.Cells(2, 6), WS2.Cells(WS2_CC_LR, 6)), 0)
'Next FOR1
'On Error GoTo 0
'
'Workbooks("得意先マスタ.xlsx").Close savechanges:=False
'
'WS2.Range("A:D").Copy
'WS2.Range("H:H").PasteSpecial
'
'End With
'
'WS2.Range(WS2.Cells(1, 8), WS2.Cells(WS2_CC_LR, 11)).RemoveDuplicates Columns:=3, Header:=xlYes
'Dim WS2_HC_LR As Long
'WS2_HC_LR = WS2.Cells(WS2.Rows.Count, 8).End(xlUp).Row
'
'    WS2.Sort.SortFields.Clear
'    WS2.Sort.SortFields.Add Key:=WS2.Range(WS2.Cells(1, 10), WS2.Cells(WS2_HC_LR, 10)) _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With WS2.Sort
'        .SetRange WS2.Range(WS2.Cells(1, 8), WS2.Cells(WS2_HC_LR, 11))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'WS2_HC_LR = WS2.Cells(WS2.Rows.Count, 8).End(xlUp).Row
'With WorksheetFunction
'For FOR1 = 2 To WS2_HC_LR
'
'WS2.Cells(FOR1, 12).Value = .SumIfs(WS2.Range("G:G"), WS2.Range("C:C"), WS2.Cells(FOR1, 10))
'Next FOR1
'
'For FOR1 = 2 To WS2_HC_LR
'
'WS2.Cells(FOR1, 13).Value = _
'.Rank(WS2.Cells(FOR1, 12), WS2.Range(WS2.Cells(2, 12), WS2.Cells(WS2_HC_LR, 12)), 0)
'Next FOR1
'End With
'
'WS2.Sort.SortFields.Clear
'    WS2.Sort.SortFields.Add Key:=WS2.Range(WS2.Cells(1, 13), WS2.Cells(WS2_HC_LR, 13)) _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With WS2.Sort
'        .SetRange WS2.Range(WS2.Cells(1, 8), WS2.Cells(WS2_HC_LR, 13))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'Dim WS1 As Worksheet
'Set WS1 = ThisWorkbook.Worksheets("得意先対比")
'
'Dim WS1_データ最終行 As Long
'WS1_データ最終行 = 5972
'
''WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
''
'    WS1.Range(WS1.Cells(3, 2), WS1.Cells(WS1_データ最終行, 3)).ClearContents
'    WS1.Range(WS1.Cells(3, 20), WS1.Cells(WS1_データ最終行, 21)).ClearContents
'
'
'
'    For FOR1 = 2 To WS2_HC_LR
'
'     WS2.Range(WS2.Cells(FOR1, 10), WS2.Cells(FOR1, 11)).Copy
'    WS1.Range(WS1.Cells(FOR1 * 5 - 7, 2), WS1.Cells(FOR1 * 5 - 3, 3)).PasteSpecial Paste:=xlPasteValues
'    WS2.Cells(FOR1, 9).Copy
'    WS1.Range(WS1.Cells(FOR1 * 5 - 7, 20), WS1.Cells(FOR1 * 5 - 3, 20)).PasteSpecial Paste:=xlPasteValues
'    WS2.Cells(FOR1, 13).Copy
'    WS1.Range(WS1.Cells(FOR1 * 5 - 7, 21), WS1.Cells(FOR1 * 5 - 3, 21)).PasteSpecial Paste:=xlPasteValues
'
'    Next FOR1
'Application.DisplayAlerts = False
'WS2.Delete
'ThisWorkbook.Save
'Application.DisplayAlerts = True
'
'WS1.Copy
'
'Set AWS1 = ActiveWorkbook.Worksheets("得意先対比")
'
'If AWS1.AutoFilterMode = True Then
'AWS1.AutoFilterMode = False
'End If
'
''ThisWorkbook.Close savechanges:=False
'AWS1.Cells.Copy
'AWS1.Cells.PasteSpecial Paste:=xlPasteValues
'
'AWS1.Range("A2:U2").AutoFilter
'
'Dim 日報_AO As String
'日報_AO = ThisWorkbook.Path & "\"
'
'Application.DisplayAlerts = False
'With WorksheetFunction
'  ActiveWorkbook.SaveAs Filename:= _
' 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_得意先別対比表.xlsx"
'End With
'
'ActiveWorkbook.Close savechanges:=False
'
''Workbooks.Open Filename:= _
''        得意先マスタ, ReadOnly:=False
''
''Set AWS1 = Workbooks("得意先マスタ.xlsx").Worksheets("sheet1")
''
''AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row
''
''        With WorksheetFunction
''
''       For FOR1 = 2 To AWS1_AC_LR
''
''        AWS1.Cells(FOR1, 1).Value = .Clean(AWS1.Cells(FOR1, 1))
''        AWS1.Cells(FOR1, 39).Value = .Clean(AWS1.Cells(FOR1, 39))
''
''        Next FOR1
''
''End With
''
''Workbooks("得意先マスタ.xlsx").Close savechanges:=True
'
'WS1.Activate
'Application.Calculation = xlCalculationAutomatic
'Application.ScreenUpdating = True
'
'MsgBox "更新完了" & vbCrLf & "処理時間は " & Round(Timer - TIME, 2) & " 秒です。"
'
'End Sub