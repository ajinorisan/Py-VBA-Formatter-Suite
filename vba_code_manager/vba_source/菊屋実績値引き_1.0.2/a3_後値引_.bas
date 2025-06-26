Option Explicit


Public Sub a3_後値引()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode = True Then
            ws.AutoFilterMode = False
        End If
    Next ws

    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Worksheets("売上取込用")
    Dim ws4 As Worksheet
    Set ws4 = ThisWorkbook.Worksheets("商品MST")
    Dim ws7 As Worksheet
    Set ws7 = ThisWorkbook.Worksheets("支払明細書")
    Dim ws8 As Worksheet
    Set ws8 = ThisWorkbook.Worksheets("加工")
    Dim ws10 As Worksheet
    Set ws10 = ThisWorkbook.Worksheets("返品用MST")
    Dim ws12 As Worksheet
    Set ws12 = ThisWorkbook.Worksheets("実績値引合計")

    Call toggle_execution_speed(False)

    ws8.Cells.ClearContents
    ws8.Cells.Style = "Normal"

    ws7.Range("D1:K1").Copy ws8.Range("A1")
    ws7.Range("P1").Copy ws8.Range("I1")

    Dim last_row_ws7 As Long
    last_row_ws7 = get_last_row(ws7, "C")
    Dim last_row_ws8 As Long
    last_row_ws8 = get_last_row(ws8, "A")



    With Application

        Dim lookup_result As Variant
        Dim i As Long
        For i = 2 To last_row_ws7
            If (ws7.Cells(i, "N") <> 0) Then

                ' データの転送
                ws8.Cells(last_row_ws8 + 1, "A").Value = ws7.Cells(i, "D").Value
                ws8.Cells(last_row_ws8 + 1, "B").Value = ws7.Cells(i, "E").Value
                ws8.Cells(last_row_ws8 + 1, "C").Value = ws7.Cells(i, "F").Value
                ws8.Cells(last_row_ws8 + 1, "D").Value = ws7.Cells(i, "G").Value
                ws8.Cells(last_row_ws8 + 1, "E").Value = ws7.Cells(i, "H").Value
                ws8.Cells(last_row_ws8 + 1, "F").Value = ws7.Cells(i, "I").Value
                ws8.Cells(last_row_ws8 + 1, "G").Value = -ws7.Cells(i, "J").Value
                ws8.Cells(last_row_ws8 + 1, "H").Value = ws7.Cells(i, "M").Value
                ws8.Cells(last_row_ws8 + 1, "I").Value = ws7.Cells(i, "P").Value
                ws8.Cells(last_row_ws8 + 1, "S").Value = Format(ws7.Cells(i, "C").Value, "yymmdd")

                ws8.Cells(last_row_ws8 + 2, "A").Value = ws7.Cells(i, "D").Value
                ws8.Cells(last_row_ws8 + 2, "B").Value = ws7.Cells(i, "E").Value
                ws8.Cells(last_row_ws8 + 2, "C").Value = ws7.Cells(i, "F").Value
                ws8.Cells(last_row_ws8 + 2, "D").Value = ws7.Cells(i, "G").Value
                ws8.Cells(last_row_ws8 + 2, "E").Value = ws7.Cells(i, "H").Value
                ws8.Cells(last_row_ws8 + 2, "F").Value = ws7.Cells(i, "I").Value
                ws8.Cells(last_row_ws8 + 2, "G").Value = ws7.Cells(i, "J").Value
                ws8.Cells(last_row_ws8 + 2, "H").Value = ws7.Cells(i, "K").Value
                ws8.Cells(last_row_ws8 + 2, "I").Value = ws7.Cells(i, "P").Value
                ws8.Cells(last_row_ws8 + 2, "S").Value = Format(ws7.Cells(i, "C").Value, "yymmdd")

                ws8.Cells(last_row_ws8 + 1, "L").Formula = ws8.Cells(last_row_ws8 + 1, "B") & ws8.Cells(last_row_ws8 + 1, "C")

                lookup_result = .VLookup(ws8.Cells(last_row_ws8 + 1, "L"), ws10.Columns("C:G"), 2, False)
                If IsError(lookup_result) Then

                    ws8.Columns.AutoFit
                    Dim last_row_ws10 As Long
                    last_row_ws10 = get_last_row(ws10, "A")
                    ws10.Cells(last_row_ws10 + 1, "A").Value = ws8.Cells(last_row_ws8 + 1, "B")
                    ws10.Cells(last_row_ws10 + 1, "B").Value = ws8.Cells(last_row_ws8 + 1, "C")
                    ws10.Cells(last_row_ws10, "C").Copy ws10.Cells(last_row_ws10 + 1, "C")

                    ws10.Activate
                    Call toggle_execution_speed(True)
                    MsgBox "返品用MSTを追加してください"
                    Exit Sub
                Else
                    ws8.Cells(last_row_ws8 + 1, "M").Value = "'" & lookup_result
                    ws8.Cells(last_row_ws8 + 1, "N").Value = .VLookup(ws8.Cells(last_row_ws8 + 1, "L"), ws10.Columns("C:G"), 3, False)
                    ws8.Cells(last_row_ws8 + 1, "O").Value = "'" & .VLookup(ws8.Cells(last_row_ws8 + 1, "L"), ws10.Columns("C:G"), 4, False)
                    ws8.Cells(last_row_ws8 + 1, "P").Value = .VLookup(ws8.Cells(last_row_ws8 + 1, "L"), ws10.Columns("C:G"), 5, False)
                End If





                ws8.Cells(last_row_ws8 + 2, "L").Formula = ws8.Cells(last_row_ws8 + 2, "B") & ws8.Cells(last_row_ws8 + 2, "C")
                ws8.Cells(last_row_ws8 + 2, "M").Value = "'" & .VLookup(ws8.Cells(last_row_ws8 + 2, "L"), ws10.Columns("C:G"), 2, False)
                ws8.Cells(last_row_ws8 + 2, "N").Value = .VLookup(ws8.Cells(last_row_ws8 + 2, "L"), ws10.Columns("C:G"), 3, False)
                ws8.Cells(last_row_ws8 + 2, "O").Value = "'" & .VLookup(ws8.Cells(last_row_ws8 + 2, "L"), ws10.Columns("C:G"), 4, False)
                ws8.Cells(last_row_ws8 + 2, "P").Value = .VLookup(ws8.Cells(last_row_ws8 + 2, "L"), ws10.Columns("C:G"), 5, False)

                ' 価格の計算

                lookup_result = .VLookup(ws8.Cells(last_row_ws8 + 1, "E") * 1, ws4.Columns("A:D"), 2, False)

                If IsError(lookup_result) Then

                    ws8.Columns.AutoFit
                    Dim last_row_ws4 As Long
                    last_row_ws4 = get_last_row(ws4, "A")
                    ws4.Cells(last_row_ws4 + 1, "A").Value = ws8.Cells(last_row_ws8 + 1, "E") * 1

                    ws4.Activate
                    Call toggle_execution_speed(True)
                    MsgBox "商品MSTを追加してください"
                    Exit Sub
                Else
                    ws8.Cells(last_row_ws8 + 1, "J").Formula = "'03"
                    ws8.Cells(last_row_ws8 + 1, "K").Value = .VLookup(ws8.Cells(last_row_ws8 + 1, "E") * 1, ws4.Columns("A:D"), 2, False)
                    ws8.Cells(last_row_ws8 + 2, "J").Formula = "'04"
                    ws8.Cells(last_row_ws8 + 2, "K").Value = .VLookup(ws8.Cells(last_row_ws8 + 2, "E") * 1, ws4.Columns("A:D"), 2, False)
                End If





                ws8.Cells(last_row_ws8 + 1, "Q").Value = ws8.Cells(last_row_ws8 + 1, "G") * ws8.Cells(last_row_ws8 + 1, "H")
                ws8.Cells(last_row_ws8 + 2, "Q").Value = ws8.Cells(last_row_ws8 + 2, "G") * ws8.Cells(last_row_ws8 + 2, "H")


                ws8.Cells(last_row_ws8 + 1, "R").Value = Mid(ws8.Cells(last_row_ws8 + 1, "M"), 3, 4)
                ws8.Cells(last_row_ws8 + 2, "R").Value = Mid(ws8.Cells(last_row_ws8 + 2, "M"), 3, 4)


                last_row_ws8 = ws8.Cells(Rows.Count, "A").End(xlUp).Row

            End If
        Next i

    End With

    ws8.Columns.AutoFit

    last_row_ws8 = ws8.Cells(Rows.Count, "A").End(xlUp).Row

    ws8.Sort.SortFields.Clear
    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, "B"), ws8.Cells(last_row_ws8, "B")), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, 13), ws8.Cells(last_row_ws8, 13)), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws8.Sort
        .SetRange ws8.Range(ws8.Cells(1, "A"), ws8.Cells(last_row_ws8, "S"))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Dim last_row_ws3 As Long
    last_row_ws3 = get_last_row(ws3, "E")
    If last_row_ws3 > 1 Then
        ws3.Rows("2:" & last_row_ws3).Delete
    End If


    last_row_ws10 = ws10.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To last_row_ws8

        ws3.Cells(i, "E").Value = ws8.Cells(i, "M")
        ws3.Cells(i, "F").Value = ws8.Cells(i, "N")
        ws3.Cells(i, "G").Value = ws8.Cells(i, "O")
        ws3.Cells(i, "H").Value = ws8.Cells(i, "P")
        ws3.Cells(i, "I").Value = "'00000998"

        ws3.Cells(i, "K").Value = ws8.Cells(i, "J")
        ws3.Cells(i, "L").Value = ws8.Cells(i, "K")

        ws3.Cells(i, "R").Value = ws8.Cells(i, "G")
        ws3.Cells(i, "S").Value = ws8.Cells(i, "H")
        ws3.Cells(i, "T").Value = ws8.Cells(i, "G") * ws8.Cells(i, "H")
        ws3.Cells(i, "U").Value = ws8.Cells(i, "S")

        ws3.Cells(i, "V").Value = Mid(ws8.Cells(i, "C"), 1, 5)

    Next i

    If WorksheetFunction.Sum(ws7.Columns("O")) <> WorksheetFunction.Sum(ws3.Columns("T")) Then

        Call toggle_execution_speed(True)
        MsgBox "後値引金額計算合いません。"
        Exit Sub
    End If

    Dim input_box As String
    input_box = InputBox("日付入力" & vbLf & vbLf & "20250301")

    If input_box = "" Then
        Call toggle_execution_speed(True)
        Exit Sub
    Else
        ws3.Cells(2, 1).Value = input_box
        last_row_ws3 = get_last_row(ws3, "E")
        ws3.Cells(2, "A").Copy ws3.Range(ws3.Cells(2, "A"), ws3.Cells(last_row_ws3, "B"))
        ws3.Cells(2, "A").Copy ws3.Range(ws3.Cells(2, "X"), ws3.Cells(last_row_ws3, "X"))
    End If

    ws3.Copy

    Application.DisplayAlerts = False

    Dim save_path As String
    Dim awb_name As String
    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋後値引"
    save_path = Environ("USERPROFILE") & "\Desktop\" & awb_name
    ActiveWorkbook.SaveAs Filename:=save_path

    ActiveWorkbook.Close

    Application.DisplayAlerts = True

    ws12.Cells.ClearContents
    ws12.Cells.Borders(xlDiagonalDown).LineStyle = xlNone
   ws12.Cells.Borders(xlDiagonalUp).LineStyle = xlNone
    ws12.Cells.Borders(xlEdgeLeft).LineStyle = xlNone
    ws12.Cells.Borders(xlEdgeTop).LineStyle = xlNone
   ws12.Cells.Borders(xlEdgeBottom).LineStyle = xlNone
    ws12.Cells.Borders(xlEdgeRight).LineStyle = xlNone
    ws12.Cells.Borders(xlInsideVertical).LineStyle = xlNone
   ws12.Cells.Borders(xlInsideHorizontal).LineStyle = xlNone

    ws8.Columns("M:N").Copy ws12.Range("A1")

    Dim last_row_ws12 As Long
    last_row_ws12 = ws12.Cells(Rows.Count, "A").End(xlUp).Row

    ws12.Range("A1:B" & last_row_ws12).RemoveDuplicates Columns:=Array(1, 2), _
                                                               Header:=xlYes

    last_row_ws12 = ws12.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To last_row_ws12
        ws12.Cells(i, "C").Value = WorksheetFunction.SumIfs(ws8.Columns("Q"), ws8.Columns("M"), ws12.Cells(i, "A"))
    Next i
    
    ws12.Cells(1, 1).Formula = "支店別後値引"
    ws12.Activate

    Call toggle_execution_speed(True)

    MsgBox "印刷してください。"

End Sub








'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws8.Sort.SortFields.Clear
'    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, 2), ws8.Cells(ws8_LR_AC, 2)), _
      '                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, 13), ws8.Cells(ws8_LR_AC, 13)), _
      '                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ws8.Sort
'        '.SetRange WS8.Range("A1:Q91")
'        .SetRange ws8.Range(ws8.Cells(1, 1), ws8.Cells(ws8_LR_AC, 18))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'
'
'    ws3.Rows("2:3000").Delete
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'    Dim ws10_LR_AC As Long
'    ws10_LR_AC = ws10.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To ws8_LR_AC
'        If ws8.Cells(i, 11) = "" Then
'            MsgBox "商品MSTを追加してください"
'            ws4.Activate
'            Exit Sub
'        ElseIf ws8.Cells(i, 13) = "" Then
'            ws10.Cells(ws10_LR_AC + 1, 1).Value = ws8.Cells(i, 2)
'            ws10.Cells(ws10_LR_AC + 1, 2).Value = ws8.Cells(i, 3)
'            MsgBox "返品用MSTを追加してください"
'            ws10.Activate
'            Exit Sub
'        Else
'            ws3.Cells(i, 5).Value = ws8.Cells(i, 13)
'            ws3.Cells(i, 6).Value = ws8.Cells(i, 14)
'            ws3.Cells(i, 7).Value = ws8.Cells(i, 15)
'            ws3.Cells(i, 8).Value = ws8.Cells(i, 16)
'            ws3.Cells(i, 9).Value = "'00000998"
'
'            'WS3.Cells(i, 10).Value = WS8.Cells(i, 1)
'            ws3.Cells(i, 11).Value = ws8.Cells(i, 10)
'            ws3.Cells(i, 12).Value = ws8.Cells(i, 11)
'
'            ws3.Cells(i, 18).Value = ws8.Cells(i, 7)
'            ws3.Cells(i, 19).Value = ws8.Cells(i, 8)
'            ws3.Cells(i, 20).Value = ws8.Cells(i, 7) * ws8.Cells(i, 8)
'            ws3.Cells(i, 21).Value = ws8.Cells(i, 19)
'
'            ws3.Cells(i, 22).Value = Mid(ws8.Cells(i, 3), 1, 5)
'
'        End If
'
'    Next i
'
'        If WorksheetFunction.Sum(ws7.columns("O")) <> WorksheetFunction.Sum(ws3.columns("T")) Then
'            MsgBox "後値引金額計算合いません。"
'            Exit Sub
'        End If
'
'
'    Dim IB As String
'    Dim ws3_LR_EC As Long
'    IB = InputBox("日付入力" & vbLf & vbLf & "20240101")
'
'    If IB = "" Then
'
'        Exit Sub
'    Else
'        ws3.Cells(2, 1).Value = IB
'        ws3_LR_EC = ws3.Cells(Rows.Count, 5).End(xlUp).Row
'        ws3.Cells(2, 1).Copy ws3.Range(ws3.Cells(2, 1), ws3.Cells(ws3_LR_EC, 2))
'        ws3.Cells(2, 1).Copy ws3.Range(ws3.Cells(2, 24), ws3.Cells(ws3_LR_EC, 24))
'    End If
'    ws3.Copy
'
'    .DisplayAlerts = False
'    Dim save_path As String
'    Dim awb_name As String
'    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋後値引"
'
'    ' 保存先とファイル名を指定（例：Desktopに保存）
'    save_path = Environ("USERPROFILE") & "\Desktop\" & awb_name
'
'    ' ワークブックを保存
'    ActiveWorkbook.SaveAs Filename:=save_path
'
'    ' ワークブックを閉じる
'    ActiveWorkbook.Close
'
'    .DisplayAlerts = True
'
'    ws8.columns("J").ClearContents
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 1 To ws8_LR_AC
'        ws8.Cells(i, 4).Value = "'" & ws8.Cells(i, 5).Value
'        ws8.Cells(i, 5).Value = ws8.Cells(i, 11).Value
'
'
'
'        If i <> 1 And (ws8.Cells(i, 18) <> ws8.Cells(i + 1, 18)) Then
'            ws8.Cells(i, 10).Value = WorksheetFunction.SumIfs(ws8.columns("Q"), ws8.columns("R"), ws8.Cells(i, 18))
'        End If
'
'
'    Next i
'
'    ws8.Range("K:S").Delete
'    ws8.Range("F:F").Delete
'
'    ws8.Range("J1").Formula = "-"
'    ws8.columns("A:P").EntireColumn.AutoFit
'    ws8.Cells.Copy WS12.Cells
'    WS12.Activate
'    Exit Sub
'
'Error:
''    Dim WS10_LR_AC As Long
'    ws10_LR_AC = ws10.Cells(Rows.Count, 1).End(xlUp).Row
'    ws10.Cells(ws10_LR_AC + 1, 1).Value = ws8.Cells(i, 2)
'    ws10.Cells(ws10_LR_AC + 1, 2).Value = ws8.Cells(i, 3)
'    ws10.Cells(ws10_LR_AC, 3).Copy ws10.Cells(ws10_LR_AC + 1, 3)
'    Resume Next
'
'syohinError:
'    Dim ws4_LR_AC As Long
'    ws4_LR_AC = ws4.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws4.Cells(ws4_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 5) * 1
'    'WS4.Cells(WS4_LR_AC + 1, 1).Value = WS8.Cells(i, 5) * 1
'    Resume Next
'End Sub

'Sub 後値引_2()
'
'    Dim ws As Worksheet
'
'    ' ワークシートをループして、オートフィルタが設定されているか確認する
'    For Each ws In ThisWorkbook.Worksheets
'        If ws.AutoFilterMode = True Then
'            ws.AutoFilterMode = False    ' オートフィルタを解除する
'        End If
'    Next ws
'
'    Dim ws3 As Worksheet
'    Set ws3 = ThisWorkbook.Worksheets("売上取込用")
'    Dim ws4 As Worksheet
'    Set ws4 = ThisWorkbook.Worksheets("商品MST")
'    Dim ws7 As Worksheet
'    Set ws7 = ThisWorkbook.Worksheets("支払明細書")
'    Dim ws8 As Worksheet
'    Set ws8 = ThisWorkbook.Worksheets("加工")
'    Dim ws10 As Worksheet
'    Set ws10 = ThisWorkbook.Worksheets("返品用MST")
'    Dim WS12 As Worksheet
'    Set WS12 = ThisWorkbook.Worksheets("実績値引合計")
'
'    ws8.Cells.ClearContents
'
'    Dim ws7_LR_CC As Long
'    ws7_LR_CC = ws7.Cells(Rows.Count, 3).End(xlUp).Row
'    Dim ws8_LR_AC As Long
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws7.Range("D1:K1").Copy ws8.Range("A1")
'    ws7.Range("P1").Copy ws8.Range("I1")
'
'    With WorksheetFunction
'
'        Dim i As Long
'        For i = 2 To ws7_LR_CC
'            If (ws7.Cells(i, 14) <> 0) Then
'
'                ws8.Cells(ws8_LR_AC + 1, 1).Value = ws7.Cells(i, 4)
'                ws8.Cells(ws8_LR_AC + 1, 2).Value = ws7.Cells(i, 5)
'                ws8.Cells(ws8_LR_AC + 1, 3).Value = ws7.Cells(i, 6)
'                ws8.Cells(ws8_LR_AC + 1, 4).Value = ws7.Cells(i, 7)
'                ws8.Cells(ws8_LR_AC + 1, 5).Value = ws7.Cells(i, 8)
'                ws8.Cells(ws8_LR_AC + 1, 6).Value = ws7.Cells(i, 9)
'                ws8.Cells(ws8_LR_AC + 1, 7).Value = -ws7.Cells(i, 10)
'                ws8.Cells(ws8_LR_AC + 1, 8).Value = ws7.Cells(i, 13)
'                ws8.Cells(ws8_LR_AC + 1, 9).Value = ws7.Cells(i, 16)
'                ws8.Cells(ws8_LR_AC + 1, 19).Value = Format(ws7.Cells(i, 3), "yymmdd")
'                ws8.Cells(ws8_LR_AC + 2, 1).Value = ws7.Cells(i, 4)
'                ws8.Cells(ws8_LR_AC + 2, 2).Value = ws7.Cells(i, 5)
'                ws8.Cells(ws8_LR_AC + 2, 3).Value = ws7.Cells(i, 6)
'                ws8.Cells(ws8_LR_AC + 2, 4).Value = ws7.Cells(i, 7)
'                ws8.Cells(ws8_LR_AC + 2, 5).Value = ws7.Cells(i, 8)
'                ws8.Cells(ws8_LR_AC + 2, 6).Value = ws7.Cells(i, 9)
'                ws8.Cells(ws8_LR_AC + 2, 7).Value = ws7.Cells(i, 10)
'                ws8.Cells(ws8_LR_AC + 2, 8).Value = ws7.Cells(i, 11)
'                ws8.Cells(ws8_LR_AC + 2, 9).Value = ws7.Cells(i, 16)
'                ws8.Cells(ws8_LR_AC + 2, 19).Value = Format(ws7.Cells(i, 3), "yymmdd")
'
'                On Error GoTo Error
'                ws8.Cells(ws8_LR_AC + 1, 12).Formula = ws8.Cells(ws8_LR_AC + 1, 2) & ws8.Cells(ws8_LR_AC + 1, 3)
'                ws8.Cells(ws8_LR_AC + 1, 13).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 1, 12), ws10.columns("C:G"), 2, False)
'                ws8.Cells(ws8_LR_AC + 1, 14).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 12), ws10.columns("C:G"), 3, False)
'                ws8.Cells(ws8_LR_AC + 1, 15).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 1, 12), ws10.columns("C:G"), 4, False)
'                ws8.Cells(ws8_LR_AC + 1, 16).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 12), ws10.columns("C:G"), 5, False)
'
'                ws8.Cells(ws8_LR_AC + 2, 12).Formula = ws8.Cells(ws8_LR_AC + 2, 2) & ws8.Cells(ws8_LR_AC + 2, 3)
'                ws8.Cells(ws8_LR_AC + 2, 13).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 2, 12), ws10.columns("C:G"), 2, False)
'                ws8.Cells(ws8_LR_AC + 2, 14).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 12), ws10.columns("C:G"), 3, False)
'                ws8.Cells(ws8_LR_AC + 2, 15).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 2, 12), ws10.columns("C:G"), 4, False)
'                ws8.Cells(ws8_LR_AC + 2, 16).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 12), ws10.columns("C:G"), 5, False)
'                On Error GoTo 0
'
'                On Error GoTo syohinError
'
'                ws8.Cells(ws8_LR_AC + 1, 10).Formula = "'03"
'                ws8.Cells(ws8_LR_AC + 1, 11).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 5) * 1, ws4.columns("A:D"), 2, False)
'                ws8.Cells(ws8_LR_AC + 2, 10).Formula = "'04"
'                ws8.Cells(ws8_LR_AC + 2, 11).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 5) * 1, ws4.columns("A:D"), 2, False)
'                On Error GoTo 0
'                ws8.Cells(ws8_LR_AC + 1, 17).Value = ws8.Cells(ws8_LR_AC + 1, 7) * ws8.Cells(ws8_LR_AC + 1, 8)
'                ws8.Cells(ws8_LR_AC + 2, 17).Value = ws8.Cells(ws8_LR_AC + 2, 7) * ws8.Cells(ws8_LR_AC + 2, 8)
'                ws8.Cells(ws8_LR_AC + 1, 18).Value = Mid(ws8.Cells(ws8_LR_AC + 1, 13), 5, 4)
'                ws8.Cells(ws8_LR_AC + 2, 18).Value = Mid(ws8.Cells(ws8_LR_AC + 2, 13), 5, 4)
'
'                ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'            End If
'        Next i
'
'    End With
'
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws8.Sort.SortFields.Clear
'    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, 2), ws8.Cells(ws8_LR_AC, 2)), _
      '                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    ws8.Sort.SortFields.Add2 Key:=ws8.Range(ws8.Cells(2, 13), ws8.Cells(ws8_LR_AC, 13)), _
      '                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With ws8.Sort
'        '.SetRange WS8.Range("A1:Q91")
'        .SetRange ws8.Range(ws8.Cells(1, 1), ws8.Cells(ws8_LR_AC, 18))
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'
'
'    ws3.Rows("2:3000").Delete
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'    Dim ws10_LR_AC As Long
'    ws10_LR_AC = ws10.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To ws8_LR_AC
'        If ws8.Cells(i, 11) = "" Then
'            MsgBox "商品MSTを追加してください"
'            ws4.Activate
'            Exit Sub
'        ElseIf ws8.Cells(i, 13) = "" Then
'            ws10.Cells(ws10_LR_AC + 1, 1).Value = ws8.Cells(i, 2)
'            ws10.Cells(ws10_LR_AC + 1, 2).Value = ws8.Cells(i, 3)
'            MsgBox "返品用MSTを追加してください"
'            ws10.Activate
'            Exit Sub
'        Else
'            ws3.Cells(i, 5).Value = ws8.Cells(i, 13)
'            ws3.Cells(i, 6).Value = ws8.Cells(i, 14)
'            ws3.Cells(i, 7).Value = ws8.Cells(i, 15)
'            ws3.Cells(i, 8).Value = ws8.Cells(i, 16)
'            ws3.Cells(i, 9).Value = "'00000998"
'
'            'WS3.Cells(i, 10).Value = WS8.Cells(i, 1)
'            ws3.Cells(i, 11).Value = ws8.Cells(i, 10)
'            ws3.Cells(i, 12).Value = ws8.Cells(i, 11)
'
'            ws3.Cells(i, 18).Value = ws8.Cells(i, 7)
'            ws3.Cells(i, 19).Value = ws8.Cells(i, 8)
'            ws3.Cells(i, 20).Value = ws8.Cells(i, 7) * ws8.Cells(i, 8)
'            ws3.Cells(i, 21).Value = ws8.Cells(i, 19)
'
'            ws3.Cells(i, 22).Value = Mid(ws8.Cells(i, 3), 1, 5)
'
'        End If
'
'    Next i
'
'    If WorksheetFunction.Sum(ws7.columns("O")) <> WorksheetFunction.Sum(ws3.columns("T")) Then
'        MsgBox "後値引金額計算合いません。"
'        Exit Sub
'    End If
'
'
'    Dim IB As String
'    Dim ws3_LR_EC As Long
'    IB = InputBox("日付入力" & vbLf & vbLf & "20240101")
'
'    If IB = "" Then
'
'        Exit Sub
'    Else
'        ws3.Cells(2, 1).Value = IB
'        ws3_LR_EC = ws3.Cells(Rows.Count, 5).End(xlUp).Row
'        ws3.Cells(2, 1).Copy ws3.Range(ws3.Cells(2, 1), ws3.Cells(ws3_LR_EC, 2))
'        ws3.Cells(2, 1).Copy ws3.Range(ws3.Cells(2, 24), ws3.Cells(ws3_LR_EC, 24))
'    End If
'    ws3.Copy
'
'    .DisplayAlerts = False
'    Dim save_path As String
'    Dim awb_name As String
'    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋後値引"
'
'    ' 保存先とファイル名を指定（例：Desktopに保存）
'    save_path = Environ("USERPROFILE") & "\Desktop\" & awb_name
'
'    ' ワークブックを保存
'    ActiveWorkbook.SaveAs Filename:=save_path
'
'    ' ワークブックを閉じる
'    ActiveWorkbook.Close
'
'    .DisplayAlerts = True
'
'    ws8.columns("J").ClearContents
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 1 To ws8_LR_AC
'        ws8.Cells(i, 4).Value = "'" & ws8.Cells(i, 5).Value
'        ws8.Cells(i, 5).Value = ws8.Cells(i, 11).Value
'
'
'
'        If i <> 1 And (ws8.Cells(i, 18) <> ws8.Cells(i + 1, 18)) Then
'            ws8.Cells(i, 10).Value = WorksheetFunction.SumIfs(ws8.columns("Q"), ws8.columns("R"), ws8.Cells(i, 18))
'        End If
'
'
'    Next i
'
'    ws8.Range("K:S").Delete
'    ws8.Range("F:F").Delete
'
'    ws8.Range("J1").Formula = "-"
'    ws8.columns("A:P").EntireColumn.AutoFit
'    ws8.Cells.Copy WS12.Cells
'    WS12.Activate
'    Exit Sub
'
'Error:
''    Dim WS10_LR_AC As Long
'    ws10_LR_AC = ws10.Cells(Rows.Count, 1).End(xlUp).Row
'    ws10.Cells(ws10_LR_AC + 1, 1).Value = ws8.Cells(i, 2)
'    ws10.Cells(ws10_LR_AC + 1, 2).Value = ws8.Cells(i, 3)
'    ws10.Cells(ws10_LR_AC, 3).Copy ws10.Cells(ws10_LR_AC + 1, 3)
'    Resume Next
'
'syohinError:
'    Dim ws4_LR_AC As Long
'    ws4_LR_AC = ws4.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws4.Cells(ws4_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 5) * 1
'    'WS4.Cells(WS4_LR_AC + 1, 1).Value = WS8.Cells(i, 5) * 1
'    Resume Next
'End Sub