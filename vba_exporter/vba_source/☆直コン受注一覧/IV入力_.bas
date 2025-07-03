Option Explicit

Sub IV入力()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    With WS1

        .Unprotect

        .Cells.EntireColumn.Hidden = False

        If .AutoFilterMode = True Then
            .AutoFilterMode = False
        End If

        Dim S1_A_LASTROW As Long    'シート1のA列の最後のデータ行取得
        S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row

        Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
        S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

        .Sort.SortFields.Clear

        Dim S1_済_COL As Long    'シート1の1行目から"済"を探す
        S1_済_COL = WorksheetFunction.Match("済", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_済_COL), .Cells(S1_A_LASTROW, S1_済_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        Dim S1_INVOICENO_COL As Long    'シート1の1行目から"INVOICE No"を探す
        S1_INVOICENO_COL = WorksheetFunction.Match("INVOICE No", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_INVOICENO_COL), .Cells(S1_A_LASTROW, S1_INVOICENO_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_IV行No_COL As Long    'シート1の1行目から"IV行No"を探す
        S1_IV行No_COL = WorksheetFunction.Match("IV行No", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_IV行No_COL), .Cells(S1_A_LASTROW, S1_IV行No_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定ETA_COL As Long    'シート1の1行目から"決定ETA"を探す
        S1_決定ETA_COL = WorksheetFunction.Match("決定ETA", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_決定ETA_COL), .Cells(S1_A_LASTROW, S1_決定ETA_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定輸入港_COL As Long    'シート1の1行目から"決定輸入港"を探す
        S1_決定輸入港_COL = WorksheetFunction.Match("決定輸入港", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_決定輸入港_COL), .Cells(S1_A_LASTROW, S1_決定輸入港_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定納品日_COL As Long    'シート1の1行目から"決定納品日"を探す
        S1_決定納品日_COL = WorksheetFunction.Match("決定納品日", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_決定納品日_COL), .Cells(S1_A_LASTROW, S1_決定納品日_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_発注書No_COL As Long    'シート1の1行目から"発注書No"を探す
        S1_発注書No_COL = WorksheetFunction.Match("発注書No", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_発注書No_COL), .Cells(S1_A_LASTROW, S1_発注書No_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_発行No_COL As Long    'シート1の1行目から"発行No"を探す
        S1_発行No_COL = WorksheetFunction.Match("発行No", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_発行No_COL), .Cells(S1_A_LASTROW, S1_発行No_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_受注番号_COL As Long    'シート1の1行目から"受注番号"を探す
        S1_受注番号_COL = WorksheetFunction.Match("受注番号", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_受注番号_COL), .Cells(S1_A_LASTROW, S1_受注番号_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_受行No_COL As Long    'シート1の1行目から"受行No"を探す
        S1_受行No_COL = WorksheetFunction.Match("受行No", .Rows("1:1"), 0)

        .Sort.SortFields.Add Key:=.Range _
        (.Cells(2, S1_受行No_COL), .Cells(S1_A_LASTROW, S1_受行No_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    End With    '①

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

        .Range(.Cells(1, 1), .Cells(1, S1_LASTCOLUMN_1)).AutoFilter Field:=S1_済_COL, _
        Criteria1:="="

        Dim S1_山善CD_COL As Long    'シート1の1行目から"山善CD"を探す
        S1_山善CD_COL = WorksheetFunction.Match("山善CD", .Rows("1:1"), 0)

        Dim S1_東洋商品名_COL As Long    'シート1の1行目から"東洋商品名"を探す
        S1_東洋商品名_COL = WorksheetFunction.Match("東洋商品名", .Rows("1:1"), 0)

        .Range(.Columns(S1_東洋商品名_COL), .Columns(S1_東洋商品名_COL)) _
        .EntireColumn.Hidden = True

        Dim S1_山善商品名_COL As Long    'シート1の1行目から"山善商品名"を探す
        S1_山善商品名_COL = WorksheetFunction.Match("山善商品名", .Rows("1:1"), 0)

        .Range(.Columns(S1_山善CD_COL), .Columns(S1_山善商品名_COL)) _
        .EntireColumn.Hidden = True

        Dim S1_希望納期_COL As Long    'シート1の1行目から"希望納期"を探す
        S1_希望納期_COL = WorksheetFunction.Match("希望納期", .Rows("1:1"), 0)

        .Range(.Columns(S1_希望納期_COL), .Columns(S1_希望納期_COL)) _
        .EntireColumn.Hidden = True

        Dim S1_卸売単価_COL As Long    'シート1の1行目から"卸売単価$"を探す
        S1_卸売単価_COL = WorksheetFunction.Match("卸売単価$", .Rows("1:1"), 0)

        Dim S1_発注日_COL As Long    'シート1の1行目から"発注日"を探す
        S1_発注日_COL = WorksheetFunction.Match("発注日", .Rows("1:1"), 0)

        .Range(.Columns(S1_卸売単価_COL), .Columns(S1_発注日_COL)) _
        .EntireColumn.Hidden = True

        Dim S1_希望ETA_COL As Long    'シート1の1行目から"希望ETA"を探す
        S1_希望ETA_COL = WorksheetFunction.Match("希望ETA", .Rows("1:1"), 0)

        Dim S1_希望輸入港_COL As Long    'シート1の1行目から"希望輸入港"を探す
        S1_希望輸入港_COL = WorksheetFunction.Match("希望輸入港", .Rows("1:1"), 0)

        '.Range(.Columns(S1_希望ETA_COL), .Columns(S1_希望輸入港_COL)) _
        '.EntireColumn.Hidden = True

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

        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True    'プロテクト
        .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能

        .Cells(1, S1_INVOICENO_COL).Activate

    End With

End Sub