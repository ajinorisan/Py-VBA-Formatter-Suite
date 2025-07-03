Option Explicit

'Sub 得意先別売上集計()
'
'Dim WS1 As Worksheet
'Set WS1 = ThisWorkbook.Worksheets("得意先対比")
'
'Dim WS2 As Worksheet
'Set WS2 = ThisWorkbook.Worksheets("貼付")
'
'Dim WS4 As Worksheet
'Set WS4 = ThisWorkbook.Worksheets("56期")
'
'Dim WS5 As Worksheet
'Set WS5 = ThisWorkbook.Worksheets("57期")
'
'Dim 売上明細 As String
'売上明細 = "C:\Users\toyocase\Desktop\日報_AO\DATA\売上明細_得意先別.xlsx"
'
'Dim 得意先マスタ As String
'得意先マスタ = "C:\Users\toyocase\Desktop\日報_AO\DATA\得意先マスタ.xlsx"
'
'Dim 日報_AO As String
'日報_AO = "C:\Users\toyocase\Desktop\日報_AO\"
'
''Dim 竜王マスタ As String
''竜王マスタ = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)\AOマスタ"
'
'Dim AWS1 As Worksheet
'Dim AWS2 As Worksheet
'
'Dim WS2_AC_LR  As Long
'Dim WS3_AC_LR As Long
'Dim WS1_AC_LR As Long
'
'Dim WS1_データ開始行 As Long
'WS1_データ開始行 = 88
'
'Dim FOR1 As Long
'
'Dim WS4_RANGE As Range
'Set WS4_RANGE = WS4.Range("A2:B729")
'
'Dim WS5_RANGE As Range
'Set WS5_RANGE = WS5.Range("A2:B712")
'
'Dim AWS1_BC_LR As Long
'Dim AWS1_SC_LR As Long
'
'If Dir(売上明細) = "" Then
'
'        MsgBox 売上明細 & vbCrLf & _
'               "が存在しません"
'
'               Exit Sub
'    Else
'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
''With Worksheets.Add()
''        .Name = "今期"
''    End With
'
'Dim WS3 As Worksheet
'Set WS3 = ThisWorkbook.Worksheets("今期")
'
'WS2.Cells.ClearContents
'
'Workbooks.Open Filename:= _
'        売上明細, ReadOnly:=True
'
'End If
'
'Set AWS1 = Workbooks("売上明細_得意先別.xlsx").Worksheets("sheet1")
'Set AWS2 = Workbooks("売上明細_得意先別.xlsx").Worksheets("sheet2")
'
'Dim AWS1_AC_LR As Long
'AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row
'Dim AWS2_AC_LR As Long
'AWS2_AC_LR = AWS2.Cells(AWS2.Rows.Count, 1).End(xlUp).Row
'
'With WorksheetFunction
'If .CountIf(AWS1.Range("E:E"), AWS2.Cells(2, 1)) = 0 Then
'AWS2.Range(AWS2.Cells(2, 1), AWS2.Cells(AWS2_AC_LR, 32)).Copy
'AWS1.Cells(AWS1_AC_LR + 1, 1).PasteSpecial Paste:=xlPasteAll
'End If
'End With
'
'AWS1.Cells.Copy
'WS2.Cells.PasteSpecial
'
'Application.CutCopyMode = False
'Workbooks("売上明細_得意先別.xlsx").Close savechanges:=False
'
'WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row
'
'With WorksheetFunction
'For FOR1 = WS2_AC_LR To 2 Step -1
'WS2.Cells(FOR1, 33).Value = Mid(WS2.Cells(FOR1, 3), 5, 2) * 1
'
'Next FOR1
'End With
'
'WS3.Cells.ClearContents
'WS2.Range(WS2.Cells(1, 1), WS2.Cells(WS2_AC_LR, 2)).Copy
'WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 2)).PasteSpecial
''WS2.Range("AG:AG").Copy
''WS3.Range("C:C").PasteSpecial
'    Application.CutCopyMode = False
'
'    WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 2)).RemoveDuplicates Columns:=1, Header:=xlYes
'    WS3.Sort.SortFields.Clear
'    WS3.Sort.SortFields.Add Key:=WS3.Range("A:A") _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With WS3.Sort
'        .SetRange WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 2))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
'
'WS4_RANGE.Copy
'WS3.Cells(WS3_AC_LR + 1, 1).PasteSpecial Paste:=xlPasteValues
'
'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
'
'WS5_RANGE.Copy
'WS3.Cells(WS3_AC_LR + 1, 1).PasteSpecial Paste:=xlPasteValues
'
'Application.CutCopyMode = False
'
'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
'
'WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 2)).RemoveDuplicates Columns:=1, Header:=xlYes
'    WS3.Sort.SortFields.Clear
'    WS3.Sort.SortFields.Add Key:=WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 1)) _
'        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With WS3.Sort
'        .SetRange WS3.Range(WS3.Cells(1, 1), WS3.Cells(WS2_AC_LR, 2))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'' WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
''WS3.Range(WS3.Cells(WS3_AC_LR, 1), WS3.Cells(WS3_AC_LR, 2)).Delete
'
'    Workbooks.Open Filename:= _
'        得意先マスタ, ReadOnly:=True
'
'Set AWS1 = Workbooks("得意先マスタ.xlsx").Worksheets("sheet1")
'
'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
'
'With WorksheetFunction
'For FOR1 = 2 To WS3_AC_LR
'WS3.Cells(FOR1, 3).Value = .VLookup(WS3.Cells(FOR1, 1), AWS1.Range("A:AN"), 39, False)
'WS3.Cells(FOR1, 4).Value = .VLookup(WS3.Cells(FOR1, 1), AWS1.Range("A:AN"), 40, False)
'WS3.Cells(FOR1, 5).Value = .SumIfs(WS2.Range("O:O"), WS2.Range("A:A"), WS3.Cells(FOR1, 1))
'Next FOR1
'
'For FOR1 = 2 To WS3_AC_LR
'WS3.Cells(FOR1, 6).Value = .Rank(WS3.Cells(FOR1, 5), WS3.Range("E:E"), 0)
'Next FOR1
''Application.DisplayAlerts = False
''Workbooks("得意先マスタ.xlsx").SaveAs 竜王マスタ & "\得意先マスタ.xlsx"
''Application.DisplayAlerts = True
'Workbooks("得意先マスタ.xlsx").Close savechanges:=False
'
'    End With
'
'
''    WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
''
''    WS1.Range(WS1.Cells(WS1_データ開始行 + 1, 2), WS1.Cells(WS1_AC_LR, 3)).ClearContents
''    WS1.Range(WS1.Cells(WS1_データ開始行 + 1, 20), WS1.Cells(WS1_AC_LR, 21)).ClearContents
''
''    For FOR1 = 1 To WS3_AC_LR
''    WS1.Range(WS1.Cells(FOR1 * 5 + WS1_データ開始行 - 4, 2), WS1.Cells(FOR1 * 5 + WS1_データ開始行, 2)) _
''    = WS3.Cells(FOR1, 1)
''    WS1.Range(WS1.Cells(FOR1 * 5 + WS1_データ開始行 - 4, 3), WS1.Cells(FOR1 * 5 + WS1_データ開始行, 3)) _
''    = WS3.Cells(FOR1, 2)
''    WS1.Range(WS1.Cells(FOR1 * 5 + WS1_データ開始行 - 4, 20), WS1.Cells(FOR1 * 5 + WS1_データ開始行, 20)) _
''    = .VLookup(WS1.Cells(FOR1 * 5 + WS1_データ開始行, 2), WS3.Range("A:F"), 4, False)
''    WS1.Range(WS1.Cells(FOR1 * 5 + WS1_データ開始行 - 4, 21), WS1.Cells(FOR1 * 5 + WS1_データ開始行, 21)) _
''    = .VLookup(WS1.Cells(FOR1 * 5 + WS1_データ開始行, 2), WS3.Range("A:F"), 6, False)
''
''    Next FOR1
''
''    End With
''
''Application.DisplayAlerts = False
''WS3.Delete
''Application.DisplayAlerts = True
''
''If WS1.AutoFilterMode = True Then
''WS1.AutoFilterMode = False
''End If
''
'WS1.Copy
'Set AWS1 = ActiveWorkbook.Worksheets("得意先対比")
''ThisWorkbook.Close savechanges:=False
'AWS1.Cells.Copy
'AWS1.Cells.PasteSpecial Paste:=xlPasteValues
''
''AWS1_BC_LR = AWS1.Cells(AWS1.Rows.Count, 2).End(xlUp).Row
''AWS1.Range(AWS1.Cells(AWS1_BC_LR + 1, 1), AWS1.Cells(10000, 21)).Delete
''
''AWS1_BC_LR = AWS1.Cells(AWS1.Rows.Count, 2).End(xlUp).Row
''
''    AWS1.Sort.SortFields.Clear
''    AWS1.Sort.SortFields.Add Key:=AWS1.Range("U" & WS1_データ開始行 + 1 & ":U" & AWS1_BC_LR) _
''        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
''    With AWS1.Sort
''        .SetRange AWS1.Range("A" & WS1_データ開始行 & ":U" & AWS1_BC_LR)
''        .Header = xlYes
''        .MatchCase = False
''        .Orientation = xlTopToBottom
''        .SortMethod = xlPinYin
''        .Apply
''    End With
''
''AWS1_BC_LR = AWS1.Cells(AWS1.Rows.Count, 2).End(xlUp).Row
''
''AWS1.Rows("3:" & WS1_データ開始行 - 1).Cut _
''Destination:=AWS1.Cells(AWS1_BC_LR + 1, 1)
''
''AWS1.Rows("2:" & WS1_データ開始行 - 1).Delete
''AWS1.Columns("A:A").Delete
''AWS1.Range("A2:T2").AutoFilter
''ActiveWindow.FreezePanes = False
''With ActiveWindow
''        .SplitColumn = 0
''        .SplitRow = 2
''    End With
''ActiveWindow.FreezePanes = True
'''ActiveWindow.FreezePanes = True
''
''AWS1_SC_LR = AWS1.Cells(AWS1.Rows.Count, 19).End(xlUp).Row
''AWS1.Rows(AWS1_SC_LR + 1 & ":10000").Delete
''
''Application.DisplayAlerts = False
''With WorksheetFunction
''  ActiveWorkbook.SaveAs Filename:= _
'' 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_得意先別対比表.xlsx"
''End With
''Application.DisplayAlerts = True
''ActiveWorkbook.Close savechanges:=False
''
'''Application.Calculation = xlCalculationAutomatic
'MsgBox "更新完了"
'WS1.Activate
'Application.ScreenUpdating = True
'
'End Sub