Option Explicit

Sub 会社別売上集計()


    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    '
    Dim TIME As Single
    '
    TIME = Timer   '時間計測開始
    '
    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("出力")

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("今期")

    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("前期")

    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("計算")

    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("変換")

    On Error Resume Next

    WS2.Visible = True
    WS3.Visible = True
    WS4.Visible = True
    WS5.Visible = True

    On Error GoTo 0

    Dim 売上明細 As String
    売上明細 = ThisWorkbook.Path & "\DATA\担当者・得意先別売上.csv"
    '
    Workbooks.Open Filename:= _
                   売上明細, ReadOnly:=True

    Dim AWS1 As Worksheet
    Set AWS1 = Workbooks("担当者・得意先別売上.csv").Worksheets(1)

    Dim AWS1_AC_LR As Long
    AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row

    AWS1.Range(AWS1_AC_LR & ":" & AWS1_AC_LR).Delete

    AWS1.Cells.Copy
    WS2.Cells.PasteSpecial

    Application.DisplayAlerts = False
    Workbooks("担当者・得意先別売上.csv").Close savechanges:=False
    Application.DisplayAlerts = True

    Dim WS2_CC_LR As Long
    WS2_CC_LR = WS2.Cells(WS2.Rows.Count, 3).End(xlUp).Row

    Dim WS3_CC_LR As Long
    WS3_CC_LR = WS3.Cells(WS3.Rows.Count, 3).End(xlUp).Row

    Dim FOR1 As Long
    For FOR1 = 2 To WS2_CC_LR
        WS2.Cells(FOR1, 7).Value = Mid(WS2.Cells(FOR1, 5), 5, 2)
    Next FOR1

    With WorksheetFunction

        WS3.Range("G:G").ClearContents

        For FOR1 = 2 To WS3_CC_LR
            If .CountIf(WS2.Range("G:G"), Mid(WS3.Cells(FOR1, 5), 5, 2)) > 0 Then
                WS3.Cells(FOR1, 7).Value = Mid(WS3.Cells(FOR1, 5), 5, 2)
            End If
        Next FOR1

    End With

    WS4.Cells.ClearContents

    If WS3.Cells(WS3_CC_LR, 1) = " 合計" Then
        WS3.Range(WS3_CC_LR & ":" & WS3_CC_LR).Delete
    End If

    WS3_CC_LR = WS3.Cells(WS3.Rows.Count, 3).End(xlUp).Row

    WS3.Range(WS3.Cells(1, 3), WS3.Cells(WS3_CC_LR, 4)).Copy
    WS4.Cells(1, 1).PasteSpecial

    WS2.Range(WS2.Cells(2, 3), WS2.Cells(WS2_CC_LR, 4)).Copy

    WS4.Cells(WS3_CC_LR + 1, 1).PasteSpecial

    Dim WS4_AC_LR As Long
    WS4_AC_LR = WS4.Cells(WS3.Rows.Count, 1).End(xlUp).Row

    Application.CutCopyMode = False

    WS4.Sort.SortFields.Clear
    WS4.Sort.SortFields.Add Key:=WS4.Range("A2:A" & WS4_AC_LR), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With WS4.Sort
        .SetRange Range("A1:B" & WS4_AC_LR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    WS4.Range("A1:B" & WS4_AC_LR).RemoveDuplicates Columns:=1, _
                                                   Header:=xlYes

    WS4.Cells(1, 3).Formula = "変換コード"
    WS4.Cells(1, 4).Formula = "変換先名称"
    WS4.Cells(1, 5).Formula = "得意先前期"
    WS4.Cells(1, 6).Formula = "得意先今期"
    WS4.Cells(1, 9).Formula = "社別前期"
    WS4.Cells(1, 10).Formula = "社別今期"

    WS4_AC_LR = WS4.Cells(WS3.Rows.Count, 1).End(xlUp).Row

    With WorksheetFunction

        For FOR1 = 2 To WS4_AC_LR
            On Error GoTo ER1
            '            Debug.Print (.Index(WS5.Cells, .Match(WS4.Cells(FOR1, 1), WS5.Range("A:A"), 0), 3))
            If .Index(WS5.Cells, .Match(WS4.Cells(FOR1, 1), WS5.Range("A:A"), 0), 4) = "" Then
ER1:
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                WS4.Activate
                WS4.Cells(FOR1, 1).Activate
                On Error GoTo 0
                MsgBox "変換マスタを作成してください"
                Exit Sub
            Else
                WS4.Cells(FOR1, 3).Value = .Index(WS5.Cells, .Match(WS4.Cells(FOR1, 1), WS5.Range("A:A"), 0), 3)
                WS4.Cells(FOR1, 4).Value = .Index(WS5.Cells, .Match(WS4.Cells(FOR1, 1), WS5.Range("A:A"), 0), 4)

            End If
        Next FOR1
        On Error GoTo 0

        For FOR1 = 2 To WS4_AC_LR

            WS4.Cells(FOR1, 5).Value = _
                                       .SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS4.Cells(FOR1, 1), WS3.Range("G:G"), "<>")
            WS4.Cells(FOR1, 6).Value = _
                                       .SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS4.Cells(FOR1, 1), WS2.Range("G:G"), "<>")

        Next FOR1

    End With

    WS4.Range("C:D").Copy
    WS4.Cells(1, 7).PasteSpecial

    Application.CutCopyMode = False

    WS4.Sort.SortFields.Clear
    WS4.Sort.SortFields.Add Key:=WS4.Range("G2:G" & WS4_AC_LR), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With WS4.Sort
        .SetRange Range("G1:H" & WS4_AC_LR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    WS4.Range("G1:H" & WS4_AC_LR).RemoveDuplicates Columns:=Array(1, 2), _
                                                   Header:=xlYes

    Dim WS4_GC_LR As Long
    WS4_GC_LR = WS4.Cells(WS3.Rows.Count, 7).End(xlUp).Row

    With WorksheetFunction
        For FOR1 = 2 To WS4_GC_LR
            WS4.Cells(FOR1, 9).Value = .SumIfs(WS4.Range("E:E"), WS4.Range("C:C"), WS4.Cells(FOR1, 7))
            WS4.Cells(FOR1, 10).Value = .SumIfs(WS4.Range("F:F"), WS4.Range("C:C"), WS4.Cells(FOR1, 7))
            WS4.Cells(FOR1, 11).Value = .Sum(WS4.Range(WS4.Cells(FOR1, 9), WS4.Cells(FOR1, 10)))
        Next FOR1
    End With

    For FOR1 = 2 To WS4_GC_LR
        If WS4.Cells(FOR1, 11) = 0 Then
            WS4.Range(WS4.Cells(FOR1, 7), WS4.Cells(FOR1, 11)).ClearContents
        End If
    Next FOR1

    WS4.Sort.SortFields.Clear
    WS4.Sort.SortFields.Add Key:=WS4.Range("G2:G" & WS4_GC_LR), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With WS4.Sort
        .SetRange Range("G1:K" & WS4_GC_LR)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    WS4.Range("K:K").Delete

    WS1.Calculate

    If WS1.AutoFilterMode = True Then
        WS1.AutoFilterMode = False
    End If
    WS1.Range("A1:I1").AutoFilter
    WS1.AutoFilter.Sort.SortFields.Clear
    WS1.AutoFilter.Sort.SortFields.Add Key:=WS1.Range( _
                                                "G1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                                       xlSortTextAsNumbers
    With WS1.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    On Error Resume Next

    WS2.Visible = False
    WS3.Visible = False
    WS4.Visible = False
    WS5.Visible = False

    On Error GoTo 0


    Application.DisplayAlerts = False

    ThisWorkbook.Save
    Application.DisplayAlerts = True

    WS1.Copy

    Set AWS1 = ActiveWorkbook.Worksheets("出力")

    AWS1.Cells.Copy
    AWS1.Cells.PasteSpecial Paste:=xlPasteValues

    AWS1.Range("A1:I1").AutoFilter Field:=7, Criteria1:="<>"


    Dim 日報_AO As String
    日報_AO = ThisWorkbook.Path & "\"

    Application.DisplayAlerts = False
    With WorksheetFunction
        ActiveWorkbook.SaveAs Filename:= _
                              日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_社別売上順位実績表.xlsx"
    End With

    ActiveWorkbook.Close savechanges:=False

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    '
    MsgBox "更新完了" & vbCrLf & "処理時間は " & Round(Timer - TIME, 2) & " 秒です。"

    'Exit Sub
    '
    'ER1:
    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    '
    'WS5.Activate
    'WS5.Cells(FOR1, 3).Activate

End Sub