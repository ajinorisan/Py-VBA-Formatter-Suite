Option Explicit

Sub 売上コンテナ並替()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    With WS1  '①

        .Unprotect    'プロテクト解除

        .Cells.EntireColumn.Hidden = False    '列の非表示解除

        If .AutoFilterMode = True Then    'オートフィルタモードなら解除
            .AutoFilterMode = False
        End If

        Dim S1_A_LASTROW As Long    'シート1のA列の最後のデータ行取得
        S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row

        Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
        S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

        .Sort.SortFields.Clear

    End With

    With WorksheetFunction

        Dim S1_済_COL As Long    'シート1の1行目から"済"を探す
        S1_済_COL = .Match("済", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_済_COL), WS1.Cells(S1_A_LASTROW, S1_済_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        Dim S1_INVOICE_COL As Long    'シート1の1行目から"INVOICE No"を探す
        S1_INVOICE_COL = .Match("INVOICE No", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_INVOICE_COL), WS1.Cells(S1_A_LASTROW, S1_INVOICE_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定納品日_COL As Long    'シート1の1行目から"決定納品日"を探す
        S1_決定納品日_COL = .Match("決定納品日", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_決定納品日_COL), WS1.Cells(S1_A_LASTROW, S1_決定納品日_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_受注番号_COL As Long    'シート1の1行目から"受注番号"を探す
        S1_受注番号_COL = .Match("受注番号", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_受注番号_COL), WS1.Cells(S1_A_LASTROW, S1_受注番号_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_受行No_COL As Long    'シート1の1行目から"受行No"を探す
        S1_受行No_COL = .Match("受行No", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_受行No_COL), WS1.Cells(S1_A_LASTROW, S1_受行No_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定ETA_COL As Long    'シート1の1行目から"決定ETA"を探す
        S1_決定ETA_COL = .Match("決定ETA", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_決定ETA_COL), WS1.Cells(S1_A_LASTROW, S1_決定ETA_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定輸入港_COL As Long    'シート1の1行目から"決定輸入港"を探す
        S1_決定輸入港_COL = .Match("決定輸入港", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_決定輸入港_COL), WS1.Cells(S1_A_LASTROW, S1_決定輸入港_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_発注書No_COL As Long    'シート1の1行目から"発注書No"を探す
        S1_発注書No_COL = .Match("発注書No", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_発注書No_COL), WS1.Cells(S1_A_LASTROW, S1_発注書No_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_発行No_COL As Long    'シート1の1行目から"発行No"を探す
        S1_発行No_COL = .Match("発行No", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
                                     (WS1.Cells(2, S1_発行No_COL), WS1.Cells(S1_A_LASTROW, S1_発行No_COL)) _
                              , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    End With

    With WS1.Sort
        .SetRange WS1.Range(WS1.Cells(1, 1), WS1.Cells(S1_A_LASTROW, S1_LASTCOLUMN_1))  '範囲をA1から右下方向は可変
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    With WS1

        .Sort.SortFields.Clear

        Dim S1_突合_LASTROW As Long

        S1_突合_LASTROW = WS1.Cells(Rows.Count, S1_済_COL + 1).End(xlUp).Row

        .Range(.Cells(2, S1_済_COL + 1), .Cells(S1_突合_LASTROW, S1_済_COL + 1)).ClearContents

        Dim FOR1 As Long
        For FOR1 = 2 To S1_A_LASTROW
            If .Cells(FOR1, S1_済_COL) = "" And .Cells(FOR1, 1) <> "" Then
                .Cells(FOR1, S1_済_COL + 1).Value = .Cells(FOR1, S1_発注書No_COL) & "-" & .Cells(FOR1, S1_発行No_COL)
            End If
        Next FOR1

        For FOR1 = 2 To S1_A_LASTROW
            If .Cells(FOR1, S1_済_COL + 1) <> "" Then
                .Cells(FOR1, S1_済_COL + 2).Value = _
                WorksheetFunction.CountIf(.Range(.Cells(2, S1_済_COL + 1), .Cells(FOR1, S1_済_COL + 1)), .Cells(FOR1, S1_済_COL + 1))
            End If
        Next FOR1

        For FOR1 = 2 To S1_A_LASTROW
            If .Cells(FOR1, S1_済_COL + 1) <> "" Then
                .Cells(FOR1, S1_済_COL + 1).Value = .Cells(FOR1, S1_済_COL + 1) & "-" & .Cells(FOR1, S1_済_COL + 2)
            End If
        Next FOR1

        .Range(.Cells(1, S1_済_COL + 2), .Cells(65536, S1_済_COL + 2)).ClearContents


        .Range(.Cells(1, 1), .Cells(1, S1_LASTCOLUMN_1)).AutoFilter Field:=S1_済_COL, _
                                                                    Criteria1:="="    '済を探してフィルタで隠す

        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
               , AllowFiltering:=True    'プロテクト

        .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能

    End With

End Sub