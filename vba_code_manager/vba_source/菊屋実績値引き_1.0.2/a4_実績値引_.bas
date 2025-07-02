Option Explicit

Public Sub a4_実績値引()

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
    Dim ws8 As Worksheet
    Set ws8 = ThisWorkbook.Worksheets("加工")
    Dim ws9 As Worksheet
    Set ws9 = ThisWorkbook.Worksheets("実績値引用MST")
    Dim ws11 As Worksheet
    Set ws11 = ThisWorkbook.Worksheets("実績値引明細書")
    Dim ws12 As Worksheet
    Set ws12 = ThisWorkbook.Worksheets("実績値引合計")

    Call toggle_execution_speed(False)

    ws8.Cells.ClearContents
    ws8.Cells.Style = "Normal"

    ws11.Range("D1:M1").Copy ws8.Range("A1")

    Dim last_row_ws11 As Long
    last_row_ws11 = get_last_row(ws11, "C")
    Dim last_row_ws8 As Long
    last_row_ws8 = get_last_row(ws8, "A")

    With Application

        Dim lookup_result As Variant

        Dim i As Long
        For i = 2 To last_row_ws11

            ws8.Cells(last_row_ws8 + 1, "A").Value = ws11.Cells(i, "D").Value
            ws8.Cells(last_row_ws8 + 1, "B").Value = ws11.Cells(i, "E").Value
            ws8.Cells(last_row_ws8 + 1, "C").Value = ws11.Cells(i, "F").Value
            ws8.Cells(last_row_ws8 + 1, "D").Value = ws11.Cells(i, "G").Value
            ws8.Cells(last_row_ws8 + 1, "E").Value = ws11.Cells(i, "H").Value
            ws8.Cells(last_row_ws8 + 1, "H").Value = -ws11.Cells(i, "K").Value
            ws8.Cells(last_row_ws8 + 1, "G").Value = ws11.Cells(i, "I").Value
            ws8.Cells(last_row_ws8 + 1, "J").Value = ws8.Cells(last_row_ws8 + 1, "G") * ws8.Cells(last_row_ws8 + 1, "H")
            ws8.Cells(last_row_ws8 + 1, "Q").Value = Format(ws11.Cells(i, "A"), "yymmdd")

            ws8.Cells(last_row_ws8 + 2, "A").Value = ws11.Cells(i, "D").Value
            ws8.Cells(last_row_ws8 + 2, "B").Value = ws11.Cells(i, "E").Value
            ws8.Cells(last_row_ws8 + 2, "C").Value = ws11.Cells(i, "F").Value
            ws8.Cells(last_row_ws8 + 2, "D").Value = ws11.Cells(i, "G").Value
            ws8.Cells(last_row_ws8 + 2, "E").Value = ws11.Cells(i, "H").Value
            ws8.Cells(last_row_ws8 + 2, "H").Value = ws11.Cells(i, "K").Value
            ws8.Cells(last_row_ws8 + 2, "G").Value = ws11.Cells(i, "J").Value
            ws8.Cells(last_row_ws8 + 2, "J").Value = ws8.Cells(last_row_ws8 + 2, "G") * ws8.Cells(last_row_ws8 + 2, "H")
            ws8.Cells(last_row_ws8 + 2, "Q").Value = Format(ws11.Cells(i, "A"), "yymmdd")

            lookup_result = .VLookup(ws8.Cells(last_row_ws8 + 1, "B") * 1, ws9.Columns("A:F"), 3, False)
            If IsError(lookup_result) Then

                Dim last_row_ws9 As Long
                last_row_ws9 = get_last_row(ws9, "A")

                ws9.Cells(last_row_ws9 + 1, "A").Value = ws8.Cells(last_row_ws8 + 1, "B").Value
                ws9.Cells(last_row_ws9 + 1, "B").Value = ws8.Cells(last_row_ws8 + 1, "C").Value
                ws9.Cells(last_row_ws9 + 1, "C").Value = "参考用：" & ws8.Cells(last_row_ws8 + 1, "A").Value

                ws9.Activate
                Call toggle_execution_speed(True)
                MsgBox "実績値引用MSTを追加してください"
                Exit Sub

            Else
                ws8.Cells(last_row_ws8 + 1, "M").Formula = "'" & lookup_result
                ws8.Cells(last_row_ws8 + 1, "N").Formula = .VLookup(ws8.Cells(last_row_ws8 + 1, "B") * 1, ws9.Columns("A:F"), 4, False)
                ws8.Cells(last_row_ws8 + 1, "O").Formula = "'" & .VLookup(ws8.Cells(last_row_ws8 + 1, "B") * 1, ws9.Columns("A:F"), 5, False)
                ws8.Cells(last_row_ws8 + 1, "P").Formula = .VLookup(ws8.Cells(last_row_ws8 + 1, "B") * 1, ws9.Columns("A:F"), 6, False)

                ws8.Cells(last_row_ws8 + 2, "M").Formula = "'" & .VLookup(ws8.Cells(last_row_ws8 + 2, "B") * 1, ws9.Columns("A:F"), 3, False)
                ws8.Cells(last_row_ws8 + 2, "N").Formula = .VLookup(ws8.Cells(last_row_ws8 + 2, "B") * 1, ws9.Columns("A:F"), 4, False)
                ws8.Cells(last_row_ws8 + 2, "O").Formula = "'" & .VLookup(ws8.Cells(last_row_ws8 + 2, "B") * 1, ws9.Columns("A:F"), 5, False)
                ws8.Cells(last_row_ws8 + 2, "P").Formula = .VLookup(ws8.Cells(last_row_ws8 + 2, "B") * 1, ws9.Columns("A:F"), 6, False)

            End If

            lookup_result = .VLookup(ws8.Cells(last_row_ws8 + 1, "D") * 1, ws4.Columns("A:D"), 2, False)
            If IsError(lookup_result) Then

                Dim last_row_ws4 As Long
                last_row_ws4 = get_last_row(ws4, "A")
                ws4.Cells(last_row_ws4 + 1, "A").Value = ws8.Cells(last_row_ws8 + 1, "D").Value
                ws4.Activate
                Call toggle_execution_speed(True)
                MsgBox "実績値引用MSTを追加してください"
                Exit Sub

            Else
                ws8.Cells(last_row_ws8 + 1, "K").Formula = "'03"
                ws8.Cells(last_row_ws8 + 1, "L").Formula = lookup_result
                ws8.Cells(last_row_ws8 + 2, "K").Formula = "'04"
                ws8.Cells(last_row_ws8 + 2, "L").Formula = .VLookup(ws8.Cells(last_row_ws8 + 2, "D") * 1, ws4.Columns("A:D"), 2, False)
            End If

            last_row_ws8 = ws8.Cells(Rows.Count, "A").End(xlUp).Row

        Next i

    End With

    ws8.Columns.AutoFit

    last_row_ws8 = ws8.Cells(Rows.Count, "A").End(xlUp).Row

    ws8.Sort.SortFields.Clear
    ws8.Sort.SortFields.Add2 Key:=Range("M2:M" & last_row_ws8), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
    With ws8.Sort
        .SetRange Range("A1:Q" & last_row_ws8)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    If WorksheetFunction.Sum(ws8.Columns("J")) <> -WorksheetFunction.Sum(ws11.Columns("M")) Then
        MsgBox "後値引金額計算合いません。"
        Exit Sub
    End If

    Dim last_row_ws3 As Long
    last_row_ws3 = get_last_row(ws3, "E")
    If last_row_ws3 > 1 Then
        ws3.Rows("2:" & last_row_ws3).Delete
    End If

    last_row_ws8 = ws8.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To last_row_ws8

        ws3.Cells(i, "E").Value = "'" & ws8.Cells(i, "M")
        ws3.Cells(i, "F").Value = ws8.Cells(i, "N")
        ws3.Cells(i, "G").Value = "'" & ws8.Cells(i, "O")
        ws3.Cells(i, "H").Value = ws8.Cells(i, "P")
        ws3.Cells(i, "I").Value = "'00000998"

        ws3.Cells(i, "K").Value = ws8.Cells(i, "K")
        ws3.Cells(i, "L").Value = ws8.Cells(i, "L")

        ws3.Cells(i, "R").Value = ws8.Cells(i, "H")
        ws3.Cells(i, "S").Value = ws8.Cells(i, "G")
        ws3.Cells(i, "T").Value = ws8.Cells(i, "J")
        ws3.Cells(i, "U").Value = ws8.Cells(i, "Q")

        ws3.Cells(i, "V").Value = Mid(ws8.Cells(i, "C"), 1, 5)

    Next i

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
    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋実績値引"
    save_path = Environ("USERPROFILE") & "\Desktop\" & awb_name
    ActiveWorkbook.SaveAs Filename:=save_path

    ActiveWorkbook.Close

    Application.DisplayAlerts = True

    ws12.Range("A:X").Delete

    ws8.Range(ws8.Cells(2, "M"), ws8.Cells(last_row_ws8, "N")).Copy ws12.Range("A2")
    ws8.Range(ws8.Cells(2, "B"), ws8.Cells(last_row_ws8, "C")).Copy ws12.Range("C2")
    ws8.Range(ws8.Cells(2, "L"), ws8.Cells(last_row_ws8, "L")).Copy ws12.Range("E2")
    ws8.Range(ws8.Cells(1, "H"), ws8.Cells(last_row_ws8, "H")).Copy ws12.Range("F1")
    ws8.Range(ws8.Cells(1, "G"), ws8.Cells(last_row_ws8, "G")).Copy ws12.Range("G1")
    ws8.Range(ws8.Cells(1, "J"), ws8.Cells(last_row_ws8, "J")).Copy ws12.Range("H1")
    ws12.Range("J1").Value = "-"

    Application.CutCopyMode = False

    Dim last_row_ws12 As Long
    last_row_ws12 = get_last_row(ws12, "A")

    For i = 2 To last_row_ws12
        If ws12.Cells(i, "A") <> ws12.Cells(i + 1, "A") Then
            ws12.Cells(i, "I").Value = WorksheetFunction.SumIfs(ws8.Range("J:J"), ws8.Range("M:M"), ws12.Cells(i, "A"))
            With ws12.Range(ws12.Cells(i + 1, "A"), ws12.Cells(i + 1, "I"))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
            End With
        End If

    Next i
    ws12.Columns("A:I").EntireColumn.AutoFit
    ws12.Activate

    Call toggle_execution_speed(True)
    MsgBox "印刷してください。"

End Sub

'    ws3.Rows("2:3000").Delete
'    For i = 2 To ws8_LR_AC
'
'        If ws8.Cells(i, 14) = "" Then
'            ws9.Activate
'            MsgBox "得意先コード登録"
'            Exit Sub
'        ElseIf ws8.Cells(i, 12) = "" Then
'            ws4.Activate
'            MsgBox "商品コード登録"
'            Exit Sub
'        Else
'            ws3.Cells(i, 5).Value = ws8.Cells(i, 13)
'            ws3.Cells(i, 6).Value = ws8.Cells(i, 14)
'            ws3.Cells(i, 7).Value = ws8.Cells(i, 15)
'            ws3.Cells(i, 8).Value = ws8.Cells(i, 16)
'            ws3.Cells(i, 9).Value = "'00000998"
'
'            'WS3.Cells(i, 10).Value = WS8.Cells(i, 1)
'            ws3.Cells(i, 11).Value = ws8.Cells(i, 11)
'            ws3.Cells(i, 12).Value = ws8.Cells(i, 12)
'
'            ws3.Cells(i, 18).Value = ws8.Cells(i, 8)
'            ws3.Cells(i, 19).Value = ws8.Cells(i, 7)
'            ws3.Cells(i, 20).Value = ws8.Cells(i, 10)
'            ws3.Cells(i, 21).Value = ws8.Cells(i, 17)
'
'            ws3.Cells(i, 22).Value = Mid(ws8.Cells(i, 3), 1, 5)
'        End If
'
'    Next i
'
'    If WorksheetFunction.Sum(ws8.columns("J")) <> WorksheetFunction.Sum(ws3.columns("T")) Then
'        MsgBox "後値引金額計算合いません。"
'        Exit Sub
'    End If
'
'    ws8.columns("A:Q").EntireColumn.AutoFit
'
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
'    Application.DisplayAlerts = False
'    Dim save_path As String
'    Dim awb_name As String
'    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋実績値引"
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
'    Application.DisplayAlerts = True
'
'
'    Dim WS12_LR_AC As Long
'
'    ws12.Range("A:X").Delete
'
'    ws8.Range(ws8.Cells(2, 13), ws8.Cells(ws8_LR_AC, 14)).Copy ws12.Range("A2")
'    Application.CutCopyMode = False
'
'    WS12_LR_AC = ws12.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws12.Range(ws12.Cells(2, 1), ws12.Cells(WS12_LR_AC, 2)).RemoveDuplicates columns:=Array(1, 2), _
'                                                                             Header:=xlNo
'
'    WS12_LR_AC = ws12.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To WS12_LR_AC
'        ws12.Cells(i, 3).Value = WorksheetFunction.SumIfs(ws8.Range("J:J"), ws8.Range("M:M"), ws12.Cells(i, 1))
'
'
'    Next i
'    ws12.columns("A:C").EntireColumn.AutoFit
'    ws12.Activate
'
'    Call toggle_execution_speed(True)
'    Exit Sub
'
'syohinError:
'    Dim ws4_LR_AC As Long
'    ws4_LR_AC = ws4.Cells(Rows.Count, 1).End(xlUp).Row
'    ' 追加する店舗用ｸﾞﾙｰﾌﾟｺｰﾄﾞ
'
'    Dim item_code As LongLong
'    item_code = ws8.Cells(ws8_LR_AC + 1, 4) * 1
'
'    Dim ITEM_RESULT As Range
'    Set ITEM_RESULT = ws4.columns("A").Find(item_code, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If ITEM_RESULT Is Nothing Then
'        ' 見つからない場合、新しい行を一番下に追加
'
'        ws4.Cells(ws4_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 4).Value
'        'WS9.Cells(WS9_LR_AC + 1, 2).Value = "商品コード登録してください。"
'
'    End If
'
'    Resume Next
'
'Error:
'    Dim WS9_LR_AC As Long
'    WS9_LR_AC = ws9.Cells(Rows.Count, 1).End(xlUp).Row
'
'
'    ' 追加する店舗用ｸﾞﾙｰﾌﾟｺｰﾄﾞ
'    Dim STORE_CODE As Long
'    STORE_CODE = ws8.Cells(ws8_LR_AC + 1, 2) * 1
'
'    Dim RESULT As Range
'    Set RESULT = ws9.columns("A").Find(STORE_CODE, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If RESULT Is Nothing Then
'        ' 見つからない場合、新しい行を一番下に追加
'
'        ws9.Cells(WS9_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 2).Value
'        ws9.Cells(WS9_LR_AC + 1, 2).Value = ws8.Cells(ws8_LR_AC + 1, 3).Value
'        ws9.Cells(WS9_LR_AC + 1, 3).Value = ws8.Cells(ws8_LR_AC + 1, 1).Value
'    End If
'
'
'
'
'    Resume Next

'    Call toggle_execution_speed(True)
'
'End Sub

'Sub 実績値引_3()
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
'    Dim ws8 As Worksheet
'    Set ws8 = ThisWorkbook.Worksheets("加工")
'    Dim ws9 As Worksheet
'    Set ws9 = ThisWorkbook.Worksheets("実績値引用MST")
'    Dim ws11 As Worksheet
'    Set ws11 = ThisWorkbook.Worksheets("実績値引明細書")
'    Dim ws12 As Worksheet
'    Set ws12 = ThisWorkbook.Worksheets("実績値引合計")
'
'    On Error Resume Next    ' エラーが発生しても続行する
'    ws11.Cells.ClearFormats
'    On Error GoTo 0    ' エラーハンドリングを元に戻す
'
'    ' セルに対する個別のスタイルの削除
'    ws11.Cells.Style = "Normal"
'
'    Dim qt As QueryTable
'    Dim FileToOpen As Variant
'
'    ws11.Cells.ClearContents
'    ' CSVファイルを開く
'    FileToOpen = Application.GetOpenFilename("実績値引明細書CSVファイル (*.csv),*.csv", , "実績値引明細書CSVファイルを選択してください")
'    If FileToOpen <> "False" Then    ' ファイルが選択された場合のみ処理を行う
'        ' クエリテーブルを作成
'        Set qt = ws11.QueryTables.Add(Connection:="TEXT;" & FileToOpen, Destination:=ws11.Range("$A$1"))
'        With qt
'            .FieldNames = True
'            .RowNumbers = False
'            .FillAdjacentFormulas = False
'            .PreserveFormatting = True
'            .RefreshOnFileOpen = False
'            .RefreshStyle = xlInsertDeleteCells
'            .SavePassword = False
'            .SaveData = True
'            .AdjustColumnWidth = True
'            .RefreshPeriod = 0
'            .TextFilePromptOnRefresh = False
'            .TextFilePlatform = 932
'            .TextFileStartRow = 1
'            .TextFileParseType = xlDelimited
'            .TextFileTextQualifier = xlTextQualifierDoubleQuote
'            .TextFileConsecutiveDelimiter = False
'            .TextFileTabDelimiter = False
'            .TextFileSemicolonDelimiter = False
'            .TextFileCommaDelimiter = True
'            .TextFileSpaceDelimiter = False
'            .TextFileColumnDataTypes = Array(1, 2, 5, 1, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 2)
'            .TextFileTrailingMinusNumbers = True
'            .Refresh BackgroundQuery:=False
'        End With
'
'        ' クエリテーブルを削除
'        For Each qt In ws11.QueryTables
'            qt.Delete
'        Next qt
'
'        ' 接続を削除
'        Dim cn As WorkbookConnection
'        For Each cn In ThisWorkbook.Connections
'            cn.Delete
'        Next cn
'
'    Else
'        Exit Sub
'    End If
'
'    ws8.Cells.ClearContents
'    ws11.Range("D1:M1").Copy ws8.Range("A1")
'
'    Dim WS11_LR_CC As Long
'    WS11_LR_CC = ws11.Cells(Rows.Count, 3).End(xlUp).Row
'    Dim ws8_LR_AC As Long
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'
'    With WorksheetFunction
'
'        Dim i As Long
'        For i = 2 To WS11_LR_CC
'
'
'            ws8.Cells(ws8_LR_AC + 1, 1).Value = ws11.Cells(i, 4)
'            ws8.Cells(ws8_LR_AC + 1, 2).Value = ws11.Cells(i, 5)
'            ws8.Cells(ws8_LR_AC + 1, 3).Value = ws11.Cells(i, 6)
'            ws8.Cells(ws8_LR_AC + 1, 4).Value = ws11.Cells(i, 7)
'            ws8.Cells(ws8_LR_AC + 1, 5).Value = ws11.Cells(i, 8)
'            ws8.Cells(ws8_LR_AC + 1, 8).Value = -ws11.Cells(i, 11)
'            ws8.Cells(ws8_LR_AC + 1, 7).Value = ws11.Cells(i, 9)
'            ws8.Cells(ws8_LR_AC + 1, 10).Value = ws8.Cells(ws8_LR_AC + 1, 7) * ws8.Cells(ws8_LR_AC + 1, 8)
'            ws8.Cells(ws8_LR_AC + 1, 17).Value = Format(ws11.Cells(i, 1), "yymmdd")
'
'            ws8.Cells(ws8_LR_AC + 2, 1).Value = ws11.Cells(i, 4)
'            ws8.Cells(ws8_LR_AC + 2, 2).Value = ws11.Cells(i, 5)
'            ws8.Cells(ws8_LR_AC + 2, 3).Value = ws11.Cells(i, 6)
'            ws8.Cells(ws8_LR_AC + 2, 4).Value = ws11.Cells(i, 7)
'            ws8.Cells(ws8_LR_AC + 2, 5).Value = ws11.Cells(i, 8)
'            ws8.Cells(ws8_LR_AC + 2, 8).Value = ws11.Cells(i, 11)
'            ws8.Cells(ws8_LR_AC + 2, 7).Value = ws11.Cells(i, 10)
'            ws8.Cells(ws8_LR_AC + 2, 10).Value = ws8.Cells(ws8_LR_AC + 2, 7) * ws8.Cells(ws8_LR_AC + 2, 8)
'            ws8.Cells(ws8_LR_AC + 2, 17).Value = Format(ws11.Cells(i, 1), "yymmdd")
'
'            On Error GoTo Error
'
'            ws8.Cells(ws8_LR_AC + 1, 13).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 1, 2) * 1, ws9.columns("A:F"), 3, False)
'            ws8.Cells(ws8_LR_AC + 1, 14).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 2) * 1, ws9.columns("A:F"), 4, False)
'            ws8.Cells(ws8_LR_AC + 1, 15).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 1, 2) * 1, ws9.columns("A:F"), 5, False)
'            ws8.Cells(ws8_LR_AC + 1, 16).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 2) * 1, ws9.columns("A:F"), 6, False)
'
'            ws8.Cells(ws8_LR_AC + 2, 13).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 2, 2) * 1, ws9.columns("A:F"), 3, False)
'            ws8.Cells(ws8_LR_AC + 2, 14).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 2) * 1, ws9.columns("A:F"), 4, False)
'            ws8.Cells(ws8_LR_AC + 2, 15).Formula = "'" & .VLookup(ws8.Cells(ws8_LR_AC + 2, 2) * 1, ws9.columns("A:F"), 5, False)
'            ws8.Cells(ws8_LR_AC + 2, 16).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 2) * 1, ws9.columns("A:F"), 6, False)
'            On Error GoTo 0
'
'            On Error GoTo syohinError
'
'            ws8.Cells(ws8_LR_AC + 1, 11).Formula = "'03"
'            ws8.Cells(ws8_LR_AC + 1, 12).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 1, 4) * 1, ws4.columns("A:D"), 2, False)
'            ws8.Cells(ws8_LR_AC + 2, 11).Formula = "'04"
'            ws8.Cells(ws8_LR_AC + 2, 12).Formula = .VLookup(ws8.Cells(ws8_LR_AC + 2, 4) * 1, ws4.columns("A:D"), 2, False)
'            On Error GoTo 0
'
'            ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'
'        Next i
'
'    End With
'
'    ws8.Sort.SortFields.Clear
'    ws8.Sort.SortFields.Add2 Key:=Range("M2:M" & ws8_LR_AC), _
'                             SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'                             xlSortTextAsNumbers
'    With ws8.Sort
'        .SetRange Range("A1:Q" & ws8_LR_AC)
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    ws3.Rows("2:3000").Delete
'    For i = 2 To ws8_LR_AC
'
'        If ws8.Cells(i, 14) = "" Then
'            ws9.Activate
'            MsgBox "得意先コード登録"
'            Exit Sub
'        ElseIf ws8.Cells(i, 12) = "" Then
'            ws4.Activate
'            MsgBox "商品コード登録"
'            Exit Sub
'        Else
'            ws3.Cells(i, 5).Value = ws8.Cells(i, 13)
'            ws3.Cells(i, 6).Value = ws8.Cells(i, 14)
'            ws3.Cells(i, 7).Value = ws8.Cells(i, 15)
'            ws3.Cells(i, 8).Value = ws8.Cells(i, 16)
'            ws3.Cells(i, 9).Value = "'00000998"
'
'            'WS3.Cells(i, 10).Value = WS8.Cells(i, 1)
'            ws3.Cells(i, 11).Value = ws8.Cells(i, 11)
'            ws3.Cells(i, 12).Value = ws8.Cells(i, 12)
'
'            ws3.Cells(i, 18).Value = ws8.Cells(i, 8)
'            ws3.Cells(i, 19).Value = ws8.Cells(i, 7)
'            ws3.Cells(i, 20).Value = ws8.Cells(i, 10)
'            ws3.Cells(i, 21).Value = ws8.Cells(i, 17)
'
'            ws3.Cells(i, 22).Value = Mid(ws8.Cells(i, 3), 1, 5)
'        End If
'
'    Next i
'
'    If WorksheetFunction.Sum(ws8.columns("J")) <> WorksheetFunction.Sum(ws3.columns("T")) Then
'        MsgBox "後値引金額計算合いません。"
'        Exit Sub
'    End If
'
'    ws8.columns("A:Q").EntireColumn.AutoFit
'
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
'    Application.DisplayAlerts = False
'    Dim save_path As String
'    Dim awb_name As String
'    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋実績値引"
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
'    Application.DisplayAlerts = True
'
'
'    Dim WS12_LR_AC As Long
'
'    ws12.Range("A:X").Delete
'
'    ws8.Range(ws8.Cells(2, 13), ws8.Cells(ws8_LR_AC, 14)).Copy ws12.Range("A2")
'    Application.CutCopyMode = False
'
'    WS12_LR_AC = ws12.Cells(Rows.Count, 1).End(xlUp).Row
'
'    ws12.Range(ws12.Cells(2, 1), ws12.Cells(WS12_LR_AC, 2)).RemoveDuplicates columns:=Array(1, 2), _
'                                                                             Header:=xlNo
'
'    WS12_LR_AC = ws12.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To WS12_LR_AC
'        ws12.Cells(i, 3).Value = WorksheetFunction.SumIfs(ws8.Range("J:J"), ws8.Range("M:M"), ws12.Cells(i, 1))
'
'
'    Next i
'    ws12.columns("A:C").EntireColumn.AutoFit
'    ws12.Activate
'    Exit Sub
'
'syohinError:
'    Dim ws4_LR_AC As Long
'    ws4_LR_AC = ws4.Cells(Rows.Count, 1).End(xlUp).Row
'    ' 追加する店舗用ｸﾞﾙｰﾌﾟｺｰﾄﾞ
'
'    Dim item_code As LongLong
'    item_code = ws8.Cells(ws8_LR_AC + 1, 4) * 1
'
'    Dim ITEM_RESULT As Range
'    Set ITEM_RESULT = ws4.columns("A").Find(item_code, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If ITEM_RESULT Is Nothing Then
'        ' 見つからない場合、新しい行を一番下に追加
'
'        ws4.Cells(ws4_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 4).Value
'        'WS9.Cells(WS9_LR_AC + 1, 2).Value = "商品コード登録してください。"
'
'    End If
'
'    Resume Next
'
'Error:
'    Dim WS9_LR_AC As Long
'    WS9_LR_AC = ws9.Cells(Rows.Count, 1).End(xlUp).Row
'
'
'    ' 追加する店舗用ｸﾞﾙｰﾌﾟｺｰﾄﾞ
'    Dim STORE_CODE As Long
'    STORE_CODE = ws8.Cells(ws8_LR_AC + 1, 2) * 1
'
'    Dim RESULT As Range
'    Set RESULT = ws9.columns("A").Find(STORE_CODE, LookIn:=xlValues, LookAt:=xlWhole)
'
'    If RESULT Is Nothing Then
'        ' 見つからない場合、新しい行を一番下に追加
'
'        ws9.Cells(WS9_LR_AC + 1, 1).Value = ws8.Cells(ws8_LR_AC + 1, 2).Value
'        ws9.Cells(WS9_LR_AC + 1, 2).Value = ws8.Cells(ws8_LR_AC + 1, 3).Value
'        ws9.Cells(WS9_LR_AC + 1, 3).Value = ws8.Cells(ws8_LR_AC + 1, 1).Value
'    End If
'
'
'
'
'    Resume Next
'
'End Sub