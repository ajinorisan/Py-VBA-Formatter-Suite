Option Explicit

Sub 直コン表作成()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    Dim WS10 As Worksheet
    Set WS10 = ThisWorkbook.Worksheets("直コン表")

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("品番マスタ")

    WS10.Cells.ClearContents
    WS10.Cells.Borders.LineStyle = False

    'Application.Calculation = xlCalculationAutomatic

    With WS1
        .Calculate
        .Unprotect

        'Application.Wait Now() + TimeValue("00:00:05")

        Dim S1_A_LASTROW As Long    'シート1のA列の最後のデータ行取得
        S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row

        Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
        S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

        Dim S1_済_COL As Long    'シート1の1行目から"済"を探す
        S1_済_COL = WorksheetFunction.Match("済", .Rows("1:1"), 0)

        .Range(.Cells(1, 1), .Cells(1, S1_LASTCOLUMN_1)).AutoFilter Field:=S1_済_COL, _
        Criteria1:="="

        Dim S1_担当者_COL As Long    'シート1の1行目から"担当者"を探す
        S1_担当者_COL = WorksheetFunction.Match("担当者", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_担当者_COL), .Cells(S1_A_LASTROW, S1_担当者_COL)).Copy
        WS10.Cells(1, 1).PasteSpecial xlValues

        Dim S1_商品CD_COL As Long    'シート1の1行目から"商品CD"を探す
        S1_商品CD_COL = WorksheetFunction.Match("商品CD", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_商品CD_COL), .Cells(S1_A_LASTROW, S1_商品CD_COL)).Copy
        WS10.Cells(1, 2).PasteSpecial xlValues

        Dim S1_山善CD_COL As Long    'シート1の1行目から"山善CD"を探す
        S1_山善CD_COL = WorksheetFunction.Match("山善CD", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_山善CD_COL), .Cells(S1_A_LASTROW, S1_山善CD_COL)).Copy
        WS10.Cells(1, 3).PasteSpecial xlValues

        Dim S1_発注日_COL As Long    'シート1の1行目から"発注日"を探す
        S1_発注日_COL = WorksheetFunction.Match("発注日", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_発注日_COL), .Cells(S1_A_LASTROW, S1_発注日_COL)).Copy
        WS10.Cells(1, 4).PasteSpecial xlValues

        Dim S1_発注書No_COL As Long    'シート1の1行目から"発注書No"を探す
        S1_発注書No_COL = WorksheetFunction.Match("発注書No", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_発注書No_COL), .Cells(S1_A_LASTROW, S1_発注書No_COL)).Copy
        WS10.Cells(1, 5).PasteSpecial xlValues

        Dim S1_発行No_COL As Long    'シート1の1行目から"発行No"を探す
        S1_発行No_COL = WorksheetFunction.Match("発行No", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_発行No_COL), .Cells(S1_A_LASTROW, S1_発行No_COL)).Copy
        WS10.Cells(1, 6).PasteSpecial xlValues

        Dim S1_受注番号_COL As Long    'シート1の1行目から"受注番号"を探す
        S1_受注番号_COL = WorksheetFunction.Match("受注番号", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_受注番号_COL), .Cells(S1_A_LASTROW, S1_受注番号_COL)).Copy
        WS10.Cells(1, 7).PasteSpecial xlValues

        Dim S1_支社_COL As Long    'シート1の1行目から"支社"を探す
        S1_支社_COL = WorksheetFunction.Match("支社", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_支社_COL), .Cells(S1_A_LASTROW, S1_支社_COL)).Copy
        WS10.Cells(1, 8).PasteSpecial xlValues

        Dim S1_数量_COL As Long    'シート1の1行目から"数量"を探す
        S1_数量_COL = WorksheetFunction.Match("数量", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_数量_COL), .Cells(S1_A_LASTROW, S1_数量_COL)).Copy
        WS10.Cells(1, 9).PasteSpecial xlValues

        Dim S1_希望輸入港_COL As Long    'シート1の1行目から"希望輸入港"を探す
        S1_希望輸入港_COL = WorksheetFunction.Match("希望輸入港", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_希望輸入港_COL), .Cells(S1_A_LASTROW, S1_希望輸入港_COL)).Copy
        WS10.Cells(1, 10).PasteSpecial xlValues

        Dim S1_入数_COL As Long    'シート1の1行目から"入数"を探す
        S1_入数_COL = WorksheetFunction.Match("入数", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_入数_COL), .Cells(S1_A_LASTROW, S1_入数_COL)).Copy
        WS10.Cells(1, 11).PasteSpecial xlValues

        Dim S1_総CBM_COL As Long    'シート1の1行目から"総CBM"を探す
        S1_総CBM_COL = WorksheetFunction.Match("総CBM", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_総CBM_COL), .Cells(S1_A_LASTROW, S1_総CBM_COL)).Copy
        WS10.Cells(1, 12).PasteSpecial xlValues

        Dim S1_決定ETA_COL As Long    'シート1の1行目から"決定ETA"を探す
        S1_決定ETA_COL = WorksheetFunction.Match("決定ETA", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_決定ETA_COL), .Cells(S1_A_LASTROW, S1_決定ETA_COL)).Copy
        WS10.Cells(1, 13).PasteSpecial xlValues

        Dim S1_決定輸入港_COL As Long    'シート1の1行目から"決定輸入港"を探す
        S1_決定輸入港_COL = WorksheetFunction.Match("決定輸入港", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_決定輸入港_COL), .Cells(S1_A_LASTROW, S1_決定輸入港_COL)).Copy
        WS10.Cells(1, 14).PasteSpecial xlValues

        Dim S1_希望ETA_COL As Long    'シート1の1行目から"希望ETA"を探す
        S1_希望ETA_COL = WorksheetFunction.Match("希望ETA", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_希望ETA_COL), .Cells(S1_A_LASTROW, S1_希望ETA_COL)).Copy
        WS10.Cells(1, 15).PasteSpecial xlValues

        Dim S1_備考_COL As Long    'シート1の1行目から"備考"を探す
        S1_備考_COL = WorksheetFunction.Match("備考", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_備考_COL), .Cells(S1_A_LASTROW, S1_備考_COL)).Copy
        WS10.Cells(1, 16).PasteSpecial xlValues

        Dim S1_発注先_COL As Long    'シート1の1行目から"発注先"を探す
        S1_発注先_COL = WorksheetFunction.Match("発注先", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_発注先_COL), .Cells(S1_A_LASTROW, S1_発注先_COL)).Copy
        WS10.Cells(1, 17).PasteSpecial xlValues

        'Dim S1_備考_COL As Long 'シート1の1行目から"INVOICE No"を探す
        'S1_備考_COL = WorksheetFunction.Match("備考", .Rows("1:1"), 0)

        .Range(.Cells(1, S1_備考_COL + 1), .Cells(S1_A_LASTROW, S1_備考_COL + 1)).Copy
        WS10.Cells(1, 18).PasteSpecial xlValues

        WS1.Calculate
        'Application.Wait Now() + TimeValue("00:00:05")

        Dim S10_AC_LR As Long
        S10_AC_LR = WS10.Cells(WS10.Rows.Count, 1).End(xlUp).Row

        Dim FOR1 As Long

        For FOR1 = 2 To S10_AC_LR
            On Error Resume Next
            'WS10.Cells(FOR1, 17).Value = WorksheetFunction.VLookup(WS10.Cells(FOR1, 3), WS2.Range("A:R"), 18, False)

            WS10.Cells(FOR1, 3).Value = WorksheetFunction.VLookup(WS10.Cells(FOR1, 3), WS2.Range("A:E"), 5, False)

            On Error GoTo 0
        Next FOR1

        WS10.Range("A:Q").EntireColumn.AutoFit

        For FOR1 = 1 To S10_AC_LR
            On Error Resume Next
            'IJRZ-3P(NV2)をIJRZ-3P(NV)に変換
            If WS10.Cells(FOR1, 2) = "IJRZ-3P(NV2)" Then
                WS10.Cells(FOR1, 2) = "IJRZ-3P(NV)"
            End If
            'IJRZ-3PF(NV2)をIJRZ-3PF(NV)に変換
            If WS10.Cells(FOR1, 2) = "IJRZ-3PF(NV2)" Then
                WS10.Cells(FOR1, 2) = "IJRZ-3PF(NV)"
            End If
            'IJRZ-1P(NV2)をIJRZ-1P(NV)に変換
            If WS10.Cells(FOR1, 2) = "IJRZ-1P(NV2)" Then
                WS10.Cells(FOR1, 2) = "IJRZ-1P(NV)"
            End If
            'IJRZ-1PF(NV2)をIJRZ-1PF(NV)に変換
            If WS10.Cells(FOR1, 2) = "IJRZ-1PF(NV2)" Then
                WS10.Cells(FOR1, 2) = "IJRZ-1PF(NV)"
                On Error GoTo 0
            End If

            If (WS10.Cells(FOR1, 13) <> WS10.Cells(FOR1 + 1, 13)) Or _
                (WS10.Cells(FOR1, 14) <> WS10.Cells(FOR1 + 1, 14)) Or _
                (WS10.Cells(FOR1, 17) <> WS10.Cells(FOR1 + 1, 17)) Then

                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeTop).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeBottom).LineStyle = xlContinuous
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

            End If
        Next FOR1

        For FOR1 = 2 To S10_AC_LR

            If (WS10.Cells(FOR1, 13) <> WS10.Cells(FOR1 - 1, 13)) Or _
                (WS10.Cells(FOR1, 14) <> WS10.Cells(FOR1 - 1, 14)) Or _
                (WS10.Cells(FOR1, 17) <> WS10.Cells(FOR1 - 1, 17)) Then

                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeTop).LineStyle = xlContinuous
                'WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 16)). _
                'Borders(xlEdgeBottom).LineStyle = xlContinuous
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                WS10.Range(WS10.Cells(FOR1, 1), WS10.Cells(FOR1, 17)). _
                Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

            End If
        Next FOR1

        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True
        .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能

        WS10.Cells(1, 3).Formula = "JAN"
        WS10.Cells(1, 17).Formula = "発注先"

        WS10.Copy
        MsgBox "作成完了"

    End With

End Sub