Option Explicit

Public Sub 月初在庫更新()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")

    Dim yes_no As Integer
    Dim yes_no2 As Integer

    yes_no = MsgBox("処理を開始しますか？", vbYesNo + vbQuestion, "確認")
    If yes_no = vbYes Then
        yes_no2 = MsgBox("月間累計の数字はコピーしましたか？" & vbCrLf & "D2セルの日付変更しましたか？", vbYesNo + vbQuestion, "確認")
        If yes_no2 = vbYes Then

            Call toggle_screen_update(False)
            Call set_calculation_mode(ws1, "manual")
            Call toggle_sheet_protection(ws1, "un_protect")
            Call set_calculation_mode(ws2, "manual")
            Call toggle_sheet_protection(ws2, "un_protect")

            Dim ws1_4_row_last_col As Long
            ws1_4_row_last_col = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

            If ws1.AutoFilterMode = True Then
                ws1.AutoFilterMode = False
                ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_4_row_last_col)).AutoFilter
            End If

            With ws1
                .Rows.Hidden = False
                .Columns.Hidden = False
            End With

            Dim ws2_4_row_last_col As Long
            ws2_4_row_last_col = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

            If ws2.AutoFilterMode = True Then
                ws2.AutoFilterMode = False
                ws2.Range(ws2.Cells(4, 1), ws2.Cells(4, ws2_4_row_last_col)).AutoFilter
            End If

            With ws2
                .Rows.Hidden = False
                .Columns.Hidden = False
            End With

            Dim day_1 As Long
            day_1 = Application.Match("月間累計", ws1.Rows(4), 0) + 1

            Dim day_last As Long
            day_last = Application.Match("翌月合計", ws1.Rows(4), 0) - 1

            ws1.Range(ws1.Cells(4, day_1), ws1.Cells(4, day_1 + 30)).ClearContents

            ws1.Cells(4, day_1).Value = ws1.Cells(2, 4)

            Dim i As Long

            For i = day_1 To day_1 + 30
                '
                If Month(ws1.Cells(4, i)) = Month(ws1.Cells(4, i) + 1) And ws1.Cells(4, i) <> "" Then

                    ws1.Cells(4, i + 1).Value = ws1.Cells(4, i) + 1

                End If

            Next i

            If Month(ws1.Cells(4, day_1 + 27)) <> 12 Then

                ws1.Cells(4, day_1 + 32).Value = Year(ws1.Cells(4, day_1 + 27)) & "/" & Month(ws1.Cells(4, day_1 + 27)) + 1 & "/1"

            Else

                ws1.Cells(4, day_1 + 32).Value = Year(ws1.Cells(4, day_1 + 27)) + 1 & "/1/1"

            End If

            day_1 = day_1 + 32

            For i = day_1 To day_last - 1
                ws1.Cells(4, i + 1).Value = ws1.Cells(4, i) + 1
            Next i

            Dim ws1_lastrow_d_col As Long
            ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

            ws1.Range(ws1.Cells(5, day_1), ws1.Cells(ws1_lastrow_d_col, day_1 + 13)).Copy



            day_1 = Application.Match("月間累計", ws1.Rows(4), 0) + 1
            ws1.Cells(5, day_1).PasteSpecial Paste:=xlPasteValues

            ws1.Range(ws1.Cells(5, day_1 + 14), ws1.Cells(ws1_lastrow_d_col - 1, day_last)).ClearContents

            Dim nyuuka_kei As Long
            nyuuka_kei = Application.Match("入荷合計", ws1.Rows(4), 0)

            Dim syukko_kei As Long
            syukko_kei = Application.Match("出庫合計", ws1.Rows(4), 0)

            ws1.Range(ws1.Cells(2, day_last + 2), ws1.Cells(ws1_lastrow_d_col - 1, nyuuka_kei - 1)).ClearContents

'            ws1.Range(ws1.Cells(2, nyuuka_kei + 2), ws1.Cells(ws1_lastrow_d_col - 1, syukko_kei - 1)).ClearContents
            
            Application.CutCopyMode = False

            Call toggle_screen_update(True)
            Call set_calculation_mode(ws1, "automatic")
            Call toggle_sheet_protection(ws1, "protect")

            Call toggle_screen_update(True)
            Call set_calculation_mode(ws2, "automatic")
            Call toggle_sheet_protection(ws2, "protect")

        End If
    End If

End Sub


'Public Sub 月初在庫更新()
'
'    Dim 月初在庫列 As Long
'    月初在庫列 = 10
'    Dim データ開始行 As Long
'    データ開始行 = 5
'    Dim 商品CD列 As Long
'    商品CD列 = 4
'
'    Dim WS4 As Worksheet
'    Set WS4 = ThisWorkbook.Worksheets("棚卸原票貼付")
'    Dim WS1 As Worksheet
'    Set WS1 = ThisWorkbook.Worksheets("在庫数")
'    Dim WS2 As Worksheet
'    Set WS2 = ThisWorkbook.Worksheets("予約")
'    Dim WS8 As Worksheet
'    Set WS8 = ThisWorkbook.Worksheets("OEM在庫")
'
'    Dim WS1_商品CD列_LR As Long
'    WS1_商品CD列_LR = WS1.Cells(WS1.Rows.Count, 商品CD列).End(xlUp).Row
'    Dim WS8_商品CD列_LR As Long
'    WS8_商品CD列_LR = WS8.Cells(WS8.Rows.Count, 商品CD列).End(xlUp).Row
'
'    Dim WS1_月間累計C_4R As Long
'    Dim WS1_28日C_4R As Long
'    Dim WS1_翌1日C_4R As Long
'
'    Dim WS8_月間累計C_4R As Long
'    Dim WS8_28日C_4R As Long
'    Dim WS8_翌1日C_4R As Long
'
'    Dim FOR1 As Long
'    Dim YESNO As Integer
'    Dim YESNO2 As Integer
'
'    YESNO = MsgBox("処理を開始しますか？", vbYesNo + vbQuestion, "確認")
'    If YESNO = vbYes Then
'        YESNO2 = MsgBox("月間累計の数字はコピーしましたか？" & vbCrLf & "D2セルの日付変更しましたか？", vbYesNo + vbQuestion, "確認")
'        If YESNO2 = vbYes Then
'
'            WS1.Unprotect
'            WS2.Unprotect
'            WS8.Unprotect
'
'            Dim WS1_LC_4R As Long
'            Dim WS2_LC_4R As Long
'            Dim WS8_LC_4R As Long
'
'            WS1_LC_4R = WS1.Cells(4, WS1.Columns.Count).End(xlToLeft).Column
'            WS2_LC_4R = WS2.Cells(4, WS2.Columns.Count).End(xlToLeft).Column
'            WS8_LC_4R = WS8.Cells(4, WS8.Columns.Count).End(xlToLeft).Column
'
'            If WS1.AutoFilterMode = True Then
'                WS1.AutoFilterMode = False
'                WS1.Range(WS1.Cells(4, 1), WS1.Cells(4, WS1_LC_4R)).AutoFilter
'            End If
'
'            If WS2.AutoFilterMode = True Then
'                WS2.AutoFilterMode = False
'                WS2.Range(WS2.Cells(4, 1), WS2.Cells(4, WS2_LC_4R)).AutoFilter
'            End If
'
'            If WS8.AutoFilterMode = True Then
'                WS8.AutoFilterMode = False
'                WS8.Range(WS8.Cells(4, 1), WS8.Cells(4, WS8_LC_4R)).AutoFilter
'            End If
'
'            With WS2
'                .Rows.Hidden = False
'                .Columns.Hidden = False
'            End With
'
'            With WS8
'                .Rows.Hidden = False
'                .Columns.Hidden = False
'            End With
'
'            With WS1
'                .Rows.Hidden = False
'                .Columns.Hidden = False
'
'            End With
'
'            With WorksheetFunction
'
'                WS1_月間累計C_4R = .Match("月間累計", WS1.Range("4:4"), 0)
'                WS1.Cells(4, WS1_月間累計C_4R + 1).Value = Year(WS1.Cells(2, 4)) & "/" & Month(WS1.Cells(2, 4)) & "/" & 1
'            End With
'
'            WS1_28日C_4R = WS1_月間累計C_4R + 28
'
'            WS1.Range(WS1.Cells(4, WS1_月間累計C_4R + 2), WS1.Cells(4, WS1_28日C_4R + 3)).ClearContents
'
'            For FOR1 = WS1_月間累計C_4R + 2 To WS1_28日C_4R
'                WS1.Cells(4, FOR1).Value = WS1.Cells(4, FOR1 - 1) + 1
'            Next FOR1
'
'            Dim DAY1 As Long
'            DAY1 = 1
'
'            For FOR1 = WS1_28日C_4R + 1 To WS1_28日C_4R + 3
'                If Month(WS1.Cells(4, WS1_28日C_4R)) = Month(WS1.Cells(4, WS1_28日C_4R) + DAY1) Then
'
'                    WS1.Cells(4, FOR1).Value = WS1.Cells(4, WS1_28日C_4R) + DAY1
'                End If
'                DAY1 = DAY1 + 1
'            Next FOR1
'            WS1_翌1日C_4R = WS1_月間累計C_4R + 33
'
'            If Month(WS1.Cells(2, 4)) = 12 Then
'                WS1.Cells(4, WS1_翌1日C_4R).Value = WS1.Cells(4, WS1_28日C_4R + 3) + 1
'            Else
'                WS1.Cells(4, WS1_翌1日C_4R).Value = Year(WS1.Cells(2, 4)) & "/" & Month(WS1.Cells(2, 4)) + 1 & "/" & 1
'            End If
'
'            For FOR1 = WS1_翌1日C_4R + 1 To WS1_翌1日C_4R + 13
'                WS1.Cells(4, FOR1).Value = WS1.Cells(4, FOR1 - 1) + 1
'            Next FOR1
'
'            With WorksheetFunction
'                WS8_月間累計C_4R = .Match("月間累計", WS8.Range("4:4"), 0)
'                WS8.Cells(4, WS8_月間累計C_4R + 1).Value = Year(WS8.Cells(2, 4)) & "/" & Month(WS8.Cells(2, 4)) & "/" & 1
'            End With
'
'            WS8_28日C_4R = WS8_月間累計C_4R + 28
'
'            WS8.Range(WS8.Cells(4, WS8_月間累計C_4R + 2), WS8.Cells(4, WS8_28日C_4R + 3)).ClearContents
'
'            For FOR1 = WS8_月間累計C_4R + 2 To WS8_28日C_4R
'                WS8.Cells(4, FOR1).Value = WS8.Cells(4, FOR1 - 1) + 1
'            Next FOR1
'
'            DAY1 = 1
'
'            For FOR1 = WS8_28日C_4R + 1 To WS8_28日C_4R + 3
'                If Month(WS8.Cells(4, WS8_28日C_4R)) = Month(WS8.Cells(4, WS8_28日C_4R) + DAY1) Then
'
'                    WS8.Cells(4, FOR1).Value = WS8.Cells(4, WS8_28日C_4R) + DAY1
'                End If
'                DAY1 = DAY1 + 1
'            Next FOR1
'            WS8_翌1日C_4R = WS8_月間累計C_4R + 33
'
'            If Month(WS8.Cells(2, 4)) = 12 Then
'                WS8.Cells(4, WS8_翌1日C_4R).Value = WS8.Cells(4, WS8_28日C_4R + 3) + 1
'            Else
'                WS8.Cells(4, WS8_翌1日C_4R).Value = Year(WS8.Cells(2, 4)) & "/" & Month(WS8.Cells(2, 4)) + 1 & "/" & 1
'            End If
'
'            For FOR1 = WS8_翌1日C_4R + 1 To WS8_翌1日C_4R + 13
'                WS8.Cells(4, FOR1).Value = WS8.Cells(4, FOR1 - 1) + 1
'            Next FOR1
'
'            WS1.Range(WS1.Cells(データ開始行, 月初在庫列), WS1.Cells(WS1_商品CD列_LR - 1, 月初在庫列)).ClearContents
'            WS8.Range(WS8.Cells(データ開始行, 月初在庫列), WS8.Cells(WS8_商品CD列_LR - 1, 月初在庫列)).ClearContents
'
'            With WorksheetFunction
'                For FOR1 = データ開始行 To WS1_商品CD列_LR - 1
'                    WS1.Cells(FOR1, 月初在庫列).Value = .SumIfs(WS4.Range("J:J"), WS4.Range("C:C"), WS1.Cells(FOR1, 商品CD列))
'                Next FOR1
'                For FOR1 = データ開始行 To WS8_商品CD列_LR - 1
'                    WS8.Cells(FOR1, 月初在庫列).Value = .SumIfs(WS4.Range("J:J"), WS4.Range("C:C"), WS8.Cells(FOR1, 商品CD列))
'                Next FOR1
'
'                '3ヶ月平均再計算
'                Dim WS1_3月平均C_4R As Long
'                Dim WS1_本日出荷C_4R As Long
'                Dim WS1_該当月C_4R As Long
'                'Dim WS1_LC_4R As Long
'                Dim WS1_DC_LR As Long
'
'                With WorksheetFunction
'                    WS1_3月平均C_4R = .Match("3ケ月平均", WS1.Range("4:4"), 0)
'                    WS1_本日出荷C_4R = .Match("本日出荷", WS1.Range("4:4"), 0)
'                    WS1_LC_4R = WS1.Cells(4, WS1.Columns.Count).End(xlToLeft).Column
'                    WS1_該当月C_4R = .Match(WS1.Cells(2, 4), WS1.Range(WS1.Cells(4, WS1_本日出荷C_4R), WS1.Cells(4, WS1_LC_4R)), 0)
'                    WS1_DC_LR = WS1.Cells(WS1.Rows.Count, 4).End(xlUp).Row
'
'                End With
'
'                WS1.Range(WS1.Cells(5, WS1_3月平均C_4R), WS1.Cells(WS1_DC_LR - 1, WS1_3月平均C_4R)).ClearContents
'
'                For FOR1 = 5 To WS1_DC_LR
'                    With WorksheetFunction
'
'                        If .CountBlank(WS1.Range(WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 4), _
                          '                                                 WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 2))) <> 3 Then
'
'                            WS1.Cells(FOR1, WS1_3月平均C_4R).Value = _
                              '                                                                  .Round(.AverageIf(WS1.Range(WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 4), _
                              '                                                                                              WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 2)), "<>""", _
                              '                                                                                    WS1.Range(WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 4), _
                              '                                                                                              WS1.Cells(FOR1, WS1_本日出荷C_4R + WS1_該当月C_4R - 2))), 0)
'
'                        Else
'                            WS1.Cells(FOR1, WS1_3月平均C_4R).Value = 0
'                        End If
'                    End With
'
'                Next FOR1
'
'
'                Dim WS8_3月平均C_4R As Long
'                Dim WS8_本日出荷C_4R As Long
'                Dim WS8_該当月C_4R As Long
'                'Dim WS8_LC_4R As Long
'                Dim WS8_DC_LR As Long
'
'                With WorksheetFunction
'                    WS8_3月平均C_4R = .Match("3ケ月平均", WS8.Range("4:4"), 0)
'                    WS8_本日出荷C_4R = .Match("本日出荷", WS8.Range("4:4"), 0)
'                    WS8_LC_4R = WS8.Cells(4, WS8.Columns.Count).End(xlToLeft).Column
'                    WS8_該当月C_4R = .Match(WS8.Cells(2, 4), WS8.Range(WS8.Cells(4, WS8_本日出荷C_4R), WS8.Cells(4, WS8_LC_4R)), 0)
'                    WS8_DC_LR = WS8.Cells(WS8.Rows.Count, 4).End(xlUp).Row
'
'                End With
'
'                WS8.Range(WS8.Cells(5, WS8_3月平均C_4R), WS8.Cells(WS8_DC_LR - 1, WS8_3月平均C_4R)).ClearContents
'
'                For FOR1 = 5 To WS8_DC_LR
'                    With WorksheetFunction
'
'                        If .CountBlank(WS8.Range(WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 4), _
                          '                                                 WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 2))) <> 3 Then
'
'                            WS8.Cells(FOR1, WS8_3月平均C_4R).Value = _
                              '                                                                  .Round(.AverageIf(WS8.Range(WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 4), _
                              '                                                                                              WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 2)), "<>""", _
                              '                                                                                    WS8.Range(WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 4), _
                              '                                                                                              WS8.Cells(FOR1, WS8_本日出荷C_4R + WS8_該当月C_4R - 2))), 0)
'
'                        Else
'                            WS8.Cells(FOR1, WS8_3月平均C_4R).Value = 0
'                        End If
'                    End With
'
'                Next FOR1
'
'                WS1.Cells(WS1_商品CD列_LR, 月初在庫列).formula = "=SUM(OFFSET(J$4,1,,ROW()-5,1))"
'                WS1.Cells(WS1_商品CD列_LR, 月初在庫列).Copy
'                WS1.Range(WS1.Cells(WS1_商品CD列_LR, 月初在庫列 + 1), WS1.Cells(WS1_商品CD列_LR, WS1_LC_4R)).PasteSpecial _
                  '                        Paste:=xlPasteFormulas
'
'                WS8.Cells(WS8_商品CD列_LR, 月初在庫列).formula = "=SUM(OFFSET(J$4,1,,ROW()-5,1))"
'                WS8.Cells(WS8_商品CD列_LR, 月初在庫列).Copy
'                WS8.Range(WS8.Cells(WS8_商品CD列_LR, 月初在庫列 + 1), WS8.Cells(WS8_商品CD列_LR, WS8_LC_4R)).PasteSpecial _
                  '                        Paste:=xlPasteFormulas
'
'                Dim WS2_引当在庫列 As Long
'                WS2_引当在庫列 = 5
'
'                'Dim WS2_LC_4R As Long
'                WS2_LC_4R = WS2.Cells(4, WS2.Columns.Count).End(xlToLeft).Column
'
'                WS2.Cells(WS1_商品CD列_LR, WS2_引当在庫列).formula = "=SUM(OFFSET(D$4,1,,ROW()-5,1))"
'                WS2.Cells(WS1_商品CD列_LR, WS2_引当在庫列).Copy
'                WS2.Range(WS2.Cells(WS1_商品CD列_LR, WS2_引当在庫列 + 1), WS2.Cells(WS1_商品CD列_LR, WS2_LC_4R)).PasteSpecial _
                  '                        Paste:=xlPasteFormulas
'
'                WS1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                  '                          , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
                  '                            AllowDeletingColumns:=True, AllowFiltering:=True
'
'                WS2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                  '                          , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
                  '                            AllowDeletingColumns:=True, AllowFiltering:=True
'
'                WS8.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                  '                          , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
                  '                            AllowDeletingColumns:=True, AllowFiltering:=True
'
'            End With
'
'        Else
'            Exit Sub
'        End If
'
'    Else
'        Exit Sub
'    End If
'
'    MsgBox "更新完了"
'
'End Sub
