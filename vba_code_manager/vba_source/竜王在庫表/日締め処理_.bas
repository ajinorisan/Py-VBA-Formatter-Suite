Option Explicit

Public Sub 日締め処理()

    Dim ws1 As Worksheet
    Dim ws4 As Worksheet
    Dim ws5 As Worksheet
    Dim ws6 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws4 = ThisWorkbook.Worksheets("入出庫明細")
    Set ws5 = ThisWorkbook.Worksheets("棚卸原票")
    Set ws6 = ThisWorkbook.Worksheets("日〆")

    Dim yes_no As Integer

    yes_no = MsgBox("日〆自動処理を開始しますか？", vbYesNo + vbQuestion, "確認")
    If yes_no = vbYes Then

        Call toggle_screen_update(False)
        Call set_calculation_mode(ws6, "manual")

        ws6.AutoFilterMode = False
        ws6.Cells.ClearContents

        Dim ws1_lastrow_e_col As Long
        ws1_lastrow_e_col = ws1.Cells(ws1.Rows.Count, "E").End(xlUp).Row

        ws1.Range(ws1.Cells(4, 4), ws1.Cells(ws1_lastrow_e_col, 5)).Copy
        ws6.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

        ws1.Range(ws1.Cells(4, 12), ws1.Cells(ws1_lastrow_e_col, 12)).Copy
        ws6.Cells(1, 3).PasteSpecial Paste:=xlPasteValues

        Dim ws6_lastrow_a_col As Long
        ws6_lastrow_a_col = ws6.Cells(ws6.Rows.Count, "A").End(xlUp).Row



        Dim date_value As Date
        date_value = DateSerial(Left(ws4.Cells(2, 1), 4), Mid(ws4.Cells(2, 1), 5, 2), Right(ws4.Cells(2, 1), 2))

        ws1.Range("$L$3").Value = Format(date_value, "yyyy/mm/dd")

        Dim honjitu_syukka As Long
        honjitu_syukka = Application.Match("本日出荷", ws1.Rows(4), 0)

        Dim syukka_day As Long
        syukka_day = Application.Match(date_value * 1, ws1.Rows(4), 0)

        ws1.Range(ws1.Cells(5, honjitu_syukka), ws1.Cells(ws1_lastrow_e_col - 1, honjitu_syukka)).Copy
        ws1.Range(ws1.Cells(5, syukka_day), ws1.Cells(ws1_lastrow_e_col - 1, syukka_day)).PasteSpecial Paste:=xlPasteValues

        ws6.Cells(1, 6).Value = "エクセル入荷"
        ws6.Cells(1, 7).Value = "AO入荷"
        ws6.Cells(1, 8).Value = "入荷差異"

        Dim nyuuka_col As Long
        nyuuka_col = Application.Match("翌月合計", ws1.Rows(4), 0) + 2

        Dim nyuuka_kei_col As Long
        nyuuka_kei_col = Application.Match("入荷合計", ws1.Rows(4), 0) - 1

        Dim i As Long
        Dim j As Long

        For i = nyuuka_col To nyuuka_kei_col

            If Not IsEmpty(ws1.Cells(3, i).Value) Then
                If CDbl(ws1.Cells(3, i).Value) = CDbl(ws1.Range("$L$3").Value) Then
                    For j = 2 To ws6_lastrow_a_col
                        ws6.Cells(j, 6).Value = ws1.Cells(j + 3, i).Value + ws6.Cells(j, 6).Value
                        ws6.Cells(j, 7).Value = Application.WorksheetFunction.SumIf(ws4.Range("G:G"), ws6.Cells(j, 1), ws4.Range("M:M")) + _
                                                Application.WorksheetFunction.SumIf(ws4.Range("G:G"), ws6.Cells(j, 1), ws4.Range("N:N"))
                        ws6.Cells(j, 8).Value = ws6.Cells(j, 6).Value - ws6.Cells(j, 7).Value
                    Next j
                End If
                If CDbl(ws1.Cells(3, i).Value) <= CDbl(ws1.Range("$L$3").Value) Then
                    For j = 2 To ws6_lastrow_a_col
                        ws6.Cells(j, 4).Value = ws1.Cells(j + 3, i).Value + ws6.Cells(j, 4).Value

                    Next j
                End If

            End If

        Next i

        ws6.Cells(1, 9).Value = "エクセル出庫"
        ws6.Cells(1, 10).Value = "AO出庫"
        ws6.Cells(1, 11).Value = "出庫差異"

        Dim syukko_col As Long
        syukko_col = Application.Match("入荷合計", ws1.Rows(4), 0) + 2

        Dim syukko_kei_col As Long
        syukko_kei_col = Application.Match("出庫合計", ws1.Rows(4), 0) - 1

        Dim found_flag As Boolean
        found_flag = False

        For i = syukko_col To syukko_kei_col

            If Not IsEmpty(ws1.Cells(3, i).Value) Then
                If CDbl(ws1.Cells(3, i).Value) = CDbl(ws1.Range("$L$3").Value) Then
                    For j = 2 To ws6_lastrow_a_col
                        ws6.Cells(j, 9).Value = ws1.Cells(j + 3, i).Value + ws6.Cells(j, 9).Value
                        ws6.Cells(j, 10).Value = Application.WorksheetFunction.SumIf(ws4.Range("G:G"), ws6.Cells(j, 1), ws4.Range("P:P"))
                        ws6.Cells(j, 11).Value = ws6.Cells(j, 9).Value - ws6.Cells(j, 10).Value
                    Next j
                End If

                If CDbl(ws1.Cells(3, i).Value) <= CDbl(ws1.Range("$L$3").Value) Then
                    For j = 2 To ws6_lastrow_a_col
                        ws6.Cells(j, 4).Value = -ws1.Cells(j + 3, i).Value + ws6.Cells(j, 4).Value
                        found_flag = True
                    Next j
                End If
            End If
        Next i
    Else
        Exit Sub
    End If

    If Not found_flag Then
        For j = 2 To ws6_lastrow_a_col
            ws6.Cells(j, 9).Value = 0
            ws6.Cells(j, 10).Value = Application.WorksheetFunction.SumIf(ws4.Range("G:G"), ws6.Cells(j, 1), ws4.Range("P:P"))
            ws6.Cells(j, 11).Value = ws6.Cells(j, 9).Value - ws6.Cells(j, 10).Value
        Next j

    End If

    ws6.Cells(1, 4).Value = "エクセル実在庫"
    ws6.Cells(1, 5).Value = "実在庫差異"
    ws6.Cells(1, 12).Value = "判定"

    Dim day_1 As Long
    day_1 = Application.Match("月間累計", ws1.Rows(4), 0) + 1

    Dim day_toujitu As Long
    day_toujitu = Application.Match(CDbl(ws1.Range("$L$3").Value), ws1.Rows(4), 0)

    Dim gessyo As Long
    gessyo = Application.Match("月初", ws1.Rows(4), 0)


    For i = 2 To ws6_lastrow_a_col

        ws6.Cells(i, 4).Value = ws1.Cells(i + 3, gessyo) - (-ws6.Cells(i, 4).Value + _
                                                            Application.WorksheetFunction.Sum(ws1.Range(ws1.Cells(i + 3, day_1), ws1.Cells(i + 3, day_toujitu))))
        ws6.Cells(i, 5).Value = ws6.Cells(i, 3).Value - ws6.Cells(i, 4).Value

        If ws6.Cells(i, 5) <> 0 Or ws6.Cells(i, 8) <> 0 Or ws6.Cells(i, 11) <> 0 Then
            ws6.Cells(i, 12).Value = "有"
        Else
            ws6.Cells(i, 12).Value = "無"
        End If

    Next i

    ws6.Range(ws6.Cells(1, 1), ws6.Cells(1, 12)).AutoFilter Field:=12, Criteria1:="有"
    ws6.Activate

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws6, "automatic")

    MsgBox "更新完了"

End Sub



'Sub 日〆処理()
'
'    Dim WS4 As Worksheet
'    Set WS4 = ThisWorkbook.Worksheets("棚卸原票貼付")
'    Dim WS1 As Worksheet
'    Set WS1 = ThisWorkbook.Worksheets("在庫数")
'    Dim WS6 As Worksheet
'    Set WS6 = ThisWorkbook.Worksheets("日〆")
'    Dim WS5 As Worksheet
'    Set WS5 = ThisWorkbook.Worksheets("入出庫明細貼付")
'
'    Dim データ開始行 As Long
'    データ開始行 = 5
'    Dim 商品CD列 As Long
'    商品CD列 = 4
'    Dim アラジン在庫列 As Long
'    アラジン在庫列 = 11
'    Dim 月間累計列 As Long
'    月間累計列 = 25
'    Dim 翌月合計列 As Long
'    翌月合計列 = 72
'    Dim 入荷開始列 As Long
'    入荷開始列 = 74
'    With WorksheetFunction
'        Dim 入荷終了列 As Long
'        入荷終了列 = .Match("入荷合計", WS1.Range("4:4"), 0)
'
'        Dim 出庫開始列 As Long
'        出庫開始列 = 入荷終了列 + 2
'        Dim 出庫終了列 As Long
'        出庫終了列 = .Match("出庫合計", WS1.Range("4:4"), 0)
'
'    End With
'
'    Dim WS1_商品CD列_LR As Long
'    WS1_商品CD列_LR = WS1.Cells(WS1.Rows.Count, 商品CD列).End(xlUp).Row
'
'    Dim WS6_AC_LR As Long
'
'    Dim FOR1 As Long
'
'    WS1.Unprotect
'
'    Dim WS1_LC_4R As Long
'
'    WS1_LC_4R = WS1.Cells(4, WS1.Columns.Count).End(xlToLeft).Column
'
'    If WS1.AutoFilterMode = True Then
'        WS1.AutoFilterMode = False
'        WS1.Range(WS1.Cells(4, 1), WS1.Cells(4, WS1_LC_4R)).AutoFilter
'    End If
'
'    With WS1
'        .Rows.Hidden = False
'        .Columns.Hidden = False
'
'    End With
'
'    If WS6.AutoFilterMode = True Then
'
'        WS6.AutoFilterMode = False
'
'    End If
'    WS6.Cells.ClearContents
'
'    WS1.Range(WS1.Cells(データ開始行 - 1, 商品CD列), WS1.Cells(WS1_商品CD列_LR - 1, 商品CD列 + 1)).Copy
'    WS6.Cells(1, 1).PasteSpecial Paste:=xlValues
'    WS1.Range(WS1.Cells(データ開始行 - 1, アラジン在庫列), WS1.Cells(WS1_商品CD列_LR - 1, アラジン在庫列)).Copy
'    WS6.Cells(1, 3).PasteSpecial Paste:=xlValues
'
'
'    WS1.Range(WS1.Cells(データ開始行 - 3, 月間累計列 + 1), WS1.Cells(データ開始行 - 3, 翌月合計列 - 1)).ClearContents
'
'    For FOR1 = 月間累計列 + 1 To 翌月合計列 - 1
'        If WS1.Cells(データ開始行 - 1, FOR1) <= WS1.Cells(3, 11) And WS1.Cells(データ開始行 - 1, FOR1) <> "" Then
'            WS1.Cells(データ開始行 - 3, FOR1).Value = 1
'        End If
'    Next FOR1
'
'    WS6_AC_LR = WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Row
'
'    'With WorksheetFunction
'    '
'    'For FOR1 = 2 To WS6_AC_LR
'    'WS6.Cells(FOR1, 4).Value = .Sum(WS1.Cells(FOR1 + 3, 10), _
      '     '-(.SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 1)))
'    'Next FOR1
'    'End With
'
'    'WS1.Range(WS1.Cells(データ開始行 - 3, 月間累計列 + 1), WS1.Cells(データ開始行 - 3, 翌月合計列 - 1)).ClearContents
'
'    WS1.Range(WS1.Cells(データ開始行 - 3, 入荷開始列), WS1.Cells(データ開始行 - 3, 入荷終了列 - 1)).ClearContents
'
'    For FOR1 = 入荷開始列 To 入荷終了列 - 1
'        If WS1.Cells(データ開始行 - 2, FOR1) <= WS1.Cells(3, 11) And WS1.Cells(データ開始行 - 2, FOR1) <> "" Then
'            WS1.Cells(データ開始行 - 3, FOR1).Value = 2
'        End If
'    Next FOR1
'
'    WS1.Range(WS1.Cells(データ開始行 - 3, 出庫開始列), WS1.Cells(データ開始行 - 3, 出庫終了列 - 1)).ClearContents
'
'    For FOR1 = 出庫開始列 To 出庫終了列 - 1
'        If WS1.Cells(データ開始行 - 2, FOR1) <= WS1.Cells(3, 11) And WS1.Cells(データ開始行 - 2, FOR1) <> "" Then
'            WS1.Cells(データ開始行 - 3, FOR1).Value = 3
'        End If
'    Next FOR1
'
'    With WorksheetFunction
'
'        For FOR1 = 2 To WS6_AC_LR
'
'            WS6.Cells(FOR1, 4).Value = .Sum(WS1.Cells(FOR1 + 3, 10), _
              '                                            -(.SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 1)), _
              '                                            -(.SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 3)), _
              '                                            .SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 2))
'
'            WS6.Cells(FOR1, 6).Value = .Sum(.SumIfs(WS5.Range("M:M"), WS5.Range("G:G"), WS6.Cells(FOR1, 1)), _
              '                                            .SumIfs(WS5.Range("N:N"), WS5.Range("G:G"), WS6.Cells(FOR1, 1)))
'
'            WS6.Cells(FOR1, 7).Value = _
              '                                       .SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 2, WS1.Range("3:3"), WS1.Cells(3, 11))
'
'            WS6.Cells(FOR1, 8).Value = .Sum(WS6.Cells(FOR1, 6), -WS6.Cells(FOR1, 7))
'
'            WS6.Cells(FOR1, 9).Value = .SumIfs(WS5.Range("P:P"), WS5.Range("G:G"), WS6.Cells(FOR1, 1))
'
'            WS6.Cells(FOR1, 10).Value = _
              '                                        .SumIfs(WS1.Range(FOR1 + 3 & ":" & FOR1 + 3), WS1.Range("2:2"), 3, WS1.Range("3:3"), WS1.Cells(3, 11))
'
'            WS6.Cells(FOR1, 11).Value = .Sum(WS6.Cells(FOR1, 9), -WS6.Cells(FOR1, 10))
'
'        Next FOR1
'
'    End With
'
'    WS1.Range(WS1.Cells(データ開始行 - 3, 月間累計列 + 1), WS1.Cells(データ開始行 - 3, 翌月合計列 - 1)).ClearContents
'    WS1.Range(WS1.Cells(データ開始行 - 3, 入荷開始列), WS1.Cells(データ開始行 - 3, 入荷終了列 - 1)).ClearContents
'    WS1.Range(WS1.Cells(データ開始行 - 3, 出庫開始列), WS1.Cells(データ開始行 - 3, 出庫終了列 - 1)).ClearContents
'
'    With WorksheetFunction
'        WS6.Cells(1, 4).formula = .text(WS1.Cells(3, 11), "m/d") & "時点実在庫"
'        WS6.Cells(1, 5).formula = "実在庫差異"
'        WS6.Cells(1, 6).formula = "アラジン入荷･入庫"
'        WS6.Cells(1, 7).formula = "在庫表入荷･入庫"
'        WS6.Cells(1, 8).formula = "入荷･入庫差異"
'        WS6.Cells(1, 9).formula = "アラジン出庫"
'        WS6.Cells(1, 10).formula = "在庫表出庫"
'        WS6.Cells(1, 11).formula = "出庫差異"
'
'        For FOR1 = 2 To WS6_AC_LR
'            WS6.Cells(FOR1, 5).Value = .Sum(WS6.Cells(FOR1, 3), -WS6.Cells(FOR1, 4))
'            WS6.Cells(FOR1, 12).Value = .Sum(WS6.Cells(FOR1, 5), WS6.Cells(FOR1, 8), WS6.Cells(FOR1, 11))
'        Next FOR1
'
'    End With
'
'    WS6.Range("A1:L1").AutoFilter Field:=12, Criteria1:="<>0"
'    WS6.Columns("A:L").EntireColumn.AutoFit
'    WS6.Activate
'
'    WS1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '              , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '                AllowDeletingColumns:=True, AllowFiltering:=True
'
'
'End Sub
