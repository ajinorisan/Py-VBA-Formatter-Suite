Sub 山善送付用明細作成()

    Dim WS1 As Worksheet
    Dim WS14 As Worksheet
    Dim WS7 As Worksheet

    Set WS1 = ThisWorkbook.Worksheets("入力")
    Set WS14 = ThisWorkbook.Worksheets("貼付")
    Set WS7 = ThisWorkbook.Worksheets("山善様送付")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call 売上コンテナ並替

    WS1.Sort.SortFields.Clear
    S1_A_LASTROW = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

    Dim S1_受注日付_COL As Long    'シート1の1行目から"受注日付"を探す
    S1_受注日付_COL = WorksheetFunction.Match("受注日付", WS1.Rows("1:1"), 0)

    WS1.Sort.SortFields.Add Key:=WS1.Range _
                                 (WS1.Cells(2, S1_受注日付_COL), WS1.Cells(S1_A_LASTROW, S1_受注日付_COL)) _
                          , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    'Dim S1_A_LASTROW As Long 'シート1のA列の最後のデータ行取得


    Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
    S1_LASTCOLUMN_1 = WS1.Cells(1, WS1.Columns.Count).End(xlToLeft).Column

    With WS1.Sort
        .SetRange WS1.Range(WS1.Cells(1, 1), WS1.Cells(S1_A_LASTROW, S1_LASTCOLUMN_1))  '範囲をA1から右下方向は可変
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    WS1.Sort.SortFields.Clear

    'Application.Calculation = xlCalculationAutomatic
    '  Application.ScreenUpdating = True
    '  AppActivate Application.Caption
    Unload UserForm1

    With WS1

        .Unprotect

        .Cells.EntireColumn.Hidden = False

        'Dim S1_A_LASTROW As Long 'シート1のA列の最後のデータ行取得
        'S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row
        '
        'Dim S1_LASTCOLUMN_1 As Long 'シート1の1行目の最後のデータ列取得
        'S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

        .Cells.Copy
        WS14.Cells.PasteSpecial xlPasteValues

        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
               , AllowFiltering:=True    'プロテクト

        .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能

    End With


    Dim WS14_A_LASTROW As Long    'シート14のA列の最後のデータ行取得
    WS14_A_LASTROW = WS14.Cells(WS14.Rows.Count, 1).End(xlUp).Row

    Dim WS7_A_LASTROW As Long    'シート7のA列の最後のデータ行取得
    WS7_A_LASTROW = WS7.Cells(WS7.Rows.Count, 1).End(xlUp).Row

    If WS7_A_LASTROW <> 1 Then

        WS7.Rows("2:" & WS7_A_LASTROW).Delete

    End If

    WS14.Range("A2:A" & WS14_A_LASTROW).Copy
    WS7.Range("A2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("G2:G" & WS14_A_LASTROW).Copy
    WS7.Range("B2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("D2:D" & WS14_A_LASTROW).Copy
    WS7.Range("C2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("B2:B" & WS14_A_LASTROW).Copy
    WS7.Range("D2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("M2:M" & WS14_A_LASTROW).Copy
    WS7.Range("E2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("F2:F" & WS14_A_LASTROW).Copy
    WS7.Range("F2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("H2:J" & WS14_A_LASTROW).Copy
    WS7.Range("G2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("L2:L" & WS14_A_LASTROW).Copy
    WS7.Range("J2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("N2:N" & WS14_A_LASTROW).Copy
    WS7.Range("K2").PasteSpecial Paste:=xlPasteValues

    WS14.Range("AK2:AM" & WS14_A_LASTROW).Copy
    WS7.Range("L2").PasteSpecial Paste:=xlPasteValues

    WS7.Range("A:N").EntireColumn.AutoFit

    WS7.Copy

End Sub