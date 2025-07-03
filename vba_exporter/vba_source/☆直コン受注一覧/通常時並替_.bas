Option Explicit

Sub 通常時並替()

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

    End With    '①

    With WorksheetFunction    '②

        Dim S1_済_COL As Long    'シート1の1行目から"済"を探す
        S1_済_COL = .Match("済", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_済_COL), WS1.Cells(S1_A_LASTROW, S1_済_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        Dim S1_決定納品日_COL As Long    'シート1の1行目から"決定納品日時"を探す
        S1_決定納品日_COL = .Match("決定納品日", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_決定納品日_COL), WS1.Cells(S1_A_LASTROW, S1_決定納品日_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_INVOICE_COL As Long    'シート1の1行目から"INVOICE No"を探す
        S1_INVOICE_COL = .Match("INVOICE No", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_INVOICE_COL), WS1.Cells(S1_A_LASTROW, S1_INVOICE_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定ETA_COL As Long    'シート1の1行目から"決定ETA"を探す
        S1_決定ETA_COL = .Match("決定ETA", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_決定ETA_COL), WS1.Cells(S1_A_LASTROW, S1_決定ETA_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_発注先_COL As Long    'シート1の1行目から"発注先"を探す
        S1_発注先_COL = .Match("発注先", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_発注先_COL), WS1.Cells(S1_A_LASTROW, S1_発注先_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_決定輸入港_COL As Long    'シート1の1行目から"決定輸入港"を探す
        S1_決定輸入港_COL = .Match("決定輸入港", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_決定輸入港_COL), WS1.Cells(S1_A_LASTROW, S1_決定輸入港_COL)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        Dim S1_希望輸入港_COL As Long    'シート1の1行目から"希望輸入港"を探す
        S1_希望輸入港_COL = .Match("希望輸入港", WS1.Rows("1:1"), 0)

        WS1.Sort.SortFields.Add Key:=WS1.Range _
        (WS1.Cells(2, S1_希望輸入港_COL), WS1.Cells(S1_A_LASTROW, S1_希望輸入港_COL)) _
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

    End With    '②

    With WS1.Sort    '③
        .SetRange WS1.Range(WS1.Cells(1, 1), WS1.Cells(S1_A_LASTROW, S1_LASTCOLUMN_1))  '範囲をA1から右下方向は可変
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With    '③

    With WS1    '④

        .Sort.SortFields.Clear

        Const S1_山善商品名_COL = 9
        Const S1_単価照合_COL = 18
        Const S1_仕入単価照会_COL = 25
        Const S1_CT数_COL = 31

        Dim S1_山善商品名_LASTROW As Long    'シート1の山善商品名列の最後のデータ行取得
        S1_山善商品名_LASTROW = .Cells(Rows.Count, S1_山善商品名_COL).End(xlUp).Row

        .Cells(2, S1_山善商品名_COL).Formula = _
        "=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH($I$1,品番マスタ!$1:$1,0))"
        .Cells(2, S1_山善商品名_COL + 1).Formula = _
        "=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH($J$1,品番マスタ!$1:$1,0))"
        .Cells(2, S1_山善商品名_COL + 2).Formula = _
        "=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH($K$1,品番マスタ!$1:$1,0))"
        .Cells(2, S1_山善商品名_COL + 3).Formula = _
        "=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH($L$1,品番マスタ!$1:$1,0))"
        .Range(.Cells(2, S1_山善商品名_COL), .Cells(2, S1_山善商品名_COL + 3)).Copy
        .Range(.Cells(3, S1_山善商品名_COL), .Cells(S1_山善商品名_LASTROW, S1_山善商品名_COL + 3)).PasteSpecial xlPasteFormulas

        .Cells(2, S1_単価照合_COL).Formula = _
        "=IF($D2=""DCM"",IF($A2<=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$P$1,品番マスタ!$1:$1,0))" & _
        ",TEXT( INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$O$1,品番マスタ!$1:$1,0)), ""#,##0.00_ ;[赤]-#,##0.00 "")" & _
        ",TEXT(INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$Q$1,品番マスタ!$1:$1,0)), ""#,##0.00_ ;[赤]-#,##0.00 ""))" & _
        ",IF($A2<=INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$P$1,品番マスタ!$1:$1,0))" & _
        ",INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$O$1,品番マスタ!$1:$1,0))," & _
        "INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$Q$1,品番マスタ!$1:$1,0))))"
        .Cells(2, S1_単価照合_COL + 1).Formula = _
        "=IF($O2<>"""",IF($AP2="""",ROUND($N2*$O2*為替!$B$1,),ROUND($AP2*$N2*$O2,)),IF($P2="""",$N2*$Q2,$N2*$P2))"
        .Range(.Cells(2, S1_単価照合_COL), .Cells(2, S1_単価照合_COL + 1)).Copy
        .Range(.Cells(3, S1_単価照合_COL), .Cells(S1_山善商品名_LASTROW, S1_単価照合_COL + 1)).PasteSpecial xlPasteFormulas

        .Cells(2, S1_仕入単価照会_COL).Formula = _
        "=IF(INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$U$1,品番マスタ!$1:$1,0))=0" & _
        ",INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$S$1,品番マスタ!$1:$1,0))," & _
        "IF(INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$T$1,品番マスタ!$1:$1,0))<=$T2," & _
        "INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$U$1,品番マスタ!$1:$1,0))" & _
        ",INDEX(品番マスタ!$1:$1048576,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$S$1,品番マスタ!$1:$1,0))))"
        .Cells(2, S1_仕入単価照会_COL + 1).Formula = _
        "=IFERROR(IF($X2="""",$Y2*$N2,$N2*$X2),0)"
        .Range(.Cells(2, S1_仕入単価照会_COL), .Cells(2, S1_仕入単価照会_COL + 1)).Copy
        .Range(.Cells(3, S1_仕入単価照会_COL), .Cells(S1_山善商品名_LASTROW, S1_仕入単価照会_COL + 1)).PasteSpecial xlPasteFormulas

        .Cells(2, S1_CT数_COL).Formula = _
        "=CEILING($N2/INDEX(品番マスタ!$A:$O,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$G$1,品番マスタ!$1:$1,0)),1)"
        .Cells(2, S1_CT数_COL + 1).Formula = _
        "=IF(ISTEXT($AE2)=TRUE,ROUND(CEILING($N2/INDEX(品番マスタ!$A:$O,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$G$1,品番マスタ!$1:$1,0)),1)" & _
        "*INDEX(品番マスタ!$A:$O,MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$L$1,品番マスタ!$1:$1,0)),3),ROUND($AE2*INDEX(品番マスタ!$A:$O," & _
        "MATCH($H2,品番マスタ!$A:$A,0),MATCH(品番マスタ!$L$1,品番マスタ!$1:$1,0)),3))"
        .Range(.Cells(2, S1_CT数_COL), .Cells(2, S1_CT数_COL + 1)).Copy
        .Range(.Cells(3, S1_CT数_COL), .Cells(S1_山善商品名_LASTROW, S1_CT数_COL)).PasteSpecial xlPasteFormulas

        WS1.Calculate

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

    End With    '④

End Sub