Option Explicit

Public Sub a2_返品処理()

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

    Call toggle_execution_speed(False)

    ws8.Cells.ClearContents
    ws8.Cells.Style = "Normal"
    ws7.Range("D1:L1").Copy ws8.Range("A1")

    Dim last_row_ws7 As Long
    last_row_ws7 = get_last_row(ws7, "C")
    Dim last_row_ws8 As Long
    last_row_ws8 = get_last_row(ws8, "A")


    Dim i As Long
    For i = 2 To last_row_ws7
        If (ws7.Cells(i, 7) = "直仕返品") Or (ws7.Cells(i, 7) = "仕入返品") Or (ws7.Cells(i, 7) = "仕入返訂") Or (ws7.Cells(i, 7) = "直仕返訂") Then

            ws8.Cells(last_row_ws8 + 1, "A").Value = ws7.Cells(i, "D").Value
            ws8.Cells(last_row_ws8 + 1, "B").Value = ws7.Cells(i, "E").Value
            ws8.Cells(last_row_ws8 + 1, "C").Value = ws7.Cells(i, "F").Value
            ws8.Cells(last_row_ws8 + 1, "D").Value = ws7.Cells(i, "G").Value
            ws8.Cells(last_row_ws8 + 1, "E").Value = ws7.Cells(i, "H").Value
            ws8.Cells(last_row_ws8 + 1, "F").Value = ws7.Cells(i, "I").Value
            ws8.Cells(last_row_ws8 + 1, "G").Value = ws7.Cells(i, "J").Value
            ws8.Cells(last_row_ws8 + 1, "H").Value = ws7.Cells(i, "K").Value
            ws8.Cells(last_row_ws8 + 1, "I").Value = ws7.Cells(i, "L").Value
            ws8.Cells(last_row_ws8 + 1, "R").Value = ws7.Cells(i, "P").Value

            last_row_ws8 = get_last_row(ws8, "A")
        End If
    Next i

    ws8.Columns.AutoFit
    
    Dim judg As Boolean
    judg = True

    With Application
        last_row_ws8 = get_last_row(ws8, "A")
        Dim error_flag As Boolean
        '        error_flag = False

        For i = 2 To last_row_ws8
            ws8.Cells(i, 10).Formula = ws8.Cells(i, "B") & ws8.Cells(i, "C")

            Dim lookup_result As Variant

            lookup_result = .VLookup(ws8.Cells(i, "J"), ws10.Columns("C:G"), 2, False)
            If IsError(lookup_result) Then

                ws8.Columns.AutoFit
                Dim last_row_ws10 As Long
                last_row_ws10 = get_last_row(ws10, "A")
                ws10.Cells(last_row_ws10 + 1, "A").Value = ws8.Cells(i, "B")
                ws10.Cells(last_row_ws10 + 1, "B").Value = ws8.Cells(i, "C")
                ws10.Cells(last_row_ws10, "C").Copy ws10.Cells(last_row_ws10 + 1, "C")
                judg = False

'                ws10.Activate
'
'
'                Call toggle_execution_speed(True)
'                MsgBox "返品用MSTを追加してください"
'                Exit Sub
            Else
                ws8.Cells(i, "K").Formula = "'" & lookup_result
            End If

            lookup_result = .VLookup(ws8.Cells(i, "J"), ws10.Columns("C:G"), 3, False)
            If Not IsError(lookup_result) Then
                ws8.Cells(i, "L").Formula = lookup_result
            End If

            lookup_result = .VLookup(ws8.Cells(i, "J"), ws10.Columns("C:G"), 4, False)
            If Not IsError(lookup_result) Then
                ws8.Cells(i, "M").Formula = "'" & lookup_result
            End If

            lookup_result = .VLookup(ws8.Cells(i, "J"), ws10.Columns("C:G"), 5, False)
            If Not IsError(lookup_result) Then
                ws8.Cells(i, "N").Formula = lookup_result
            End If

            Dim last_row_ws4 As Long

            If Len(ws8.Cells(i, "E")) < 13 Then
                ws8.Cells(i, "O").Formula = "'03"
                lookup_result = .VLookup(ws8.Cells(i, "E") * 1, ws4.Columns("A:D"), 2, False)
                If IsError(lookup_result) Then

                    ws8.Columns.AutoFit
                    last_row_ws4 = get_last_row(ws4, "A")
                    ws4.Cells(last_row_ws4 + 1, "A").Value = ws8.Cells(i, "E") * 1
                    judg = False
'                    ws4.Activate
'                    Call toggle_execution_speed(True)
'                    MsgBox "商品MSTを追加してください"
'
'                    Exit Sub
                Else
                    ws8.Cells(i, "P").Formula = lookup_result
                End If

                lookup_result = .VLookup(ws8.Cells(i, "E") * 1, ws4.Columns("A:D"), 3, False)
                If Not IsError(lookup_result) Then
                    ws8.Cells(i, "Q").Formula = lookup_result
                End If
            Else
                ws8.Cells(i, "O").Formula = "'02"
                lookup_result = .VLookup(ws8.Cells(i, "E") * 1, ws4.Columns("A:D"), 2, False)
                If IsError(lookup_result) Then
                    ws8.Columns.AutoFit
                    last_row_ws4 = get_last_row(ws4, "A")
                    ws4.Cells(last_row_ws4 + 1, "A").Value = ws8.Cells(i, "E") * 1
                    
                    judg = False
                    
'                    ws4.Activate
'                    Call toggle_execution_speed(True)
'                    MsgBox "商品MSTを追加してください"
'
'                    Exit Sub
                Else
                    ws8.Cells(i, "P").Formula = lookup_result
                End If
            End If
        Next i

    End With

    ws8.Columns.AutoFit
    
    If judg = False Then
    ws10.Activate

                Call toggle_execution_speed(True)
                MsgBox "返品用MST or 商品MSTを追加してください"
                Exit Sub
    End If

    Dim last_row_ws3 As Long
    last_row_ws3 = get_last_row(ws3, "E")
    If last_row_ws3 > 1 Then
        ws3.Rows("2:" & last_row_ws3).Delete
    End If

    last_row_ws8 = get_last_row(ws8, "A")

    For i = 2 To last_row_ws8
        ws3.Cells(i, "E").Value = "'" & ws8.Cells(i, "K").Value
        ws3.Cells(i, "F").Value = ws8.Cells(i, "L").Value
        ws3.Cells(i, "G").Value = "'" & ws8.Cells(i, "M").Value
        ws3.Cells(i, "H").Value = ws8.Cells(i, "N").Value
        If ws8.Cells(i, "O").Value = "02" Then
            ws3.Cells(i, "I").Value = "'00000999"
        Else
            ws3.Cells(i, "I").Value = "'00000998"
        End If
        ws3.Cells(i, "J").Value = ws8.Cells(i, "A").Value
        ws3.Cells(i, "K").Value = "'" & ws8.Cells(i, "O").Value
        ws3.Cells(i, "L").Value = ws8.Cells(i, "P").Value
        If Len(ws8.Cells(i, "E")) < 13 Then
            ws3.Cells(i, "N").Value = ws8.Cells(i, "Q").Value
        End If
            ws3.Cells(i, "Y").Value = Format(ws8.Cells(i, "R").Value, "yyyymm")
        
        ws3.Cells(i, "R").Value = ws8.Cells(i, "G").Value
        ws3.Cells(i, "S").Value = ws8.Cells(i, "H").Value
        ws3.Cells(i, "T").Value = ws8.Cells(i, "I").Value
    Next i

    If WorksheetFunction.Sum(ws8.Columns("I")) <> WorksheetFunction.Sum(ws3.Columns("T")) Then

        Call toggle_execution_speed(True)
        MsgBox "返品金額計算合いません。"

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
    awb_name = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋返品"
    save_path = Environ("USERPROFILE") & "\Desktop\" & awb_name
    ActiveWorkbook.SaveAs Filename:=save_path

    Application.DisplayAlerts = True

    Call toggle_execution_speed(True)

    MsgBox "既に計上済みの伝票を削除してください"

    'ActiveWorkbook.Close

End Sub

'
'
'    With WorksheetFunction
'
'        last_row_ws8 = get_last_row(ws8, "A")
'        For i = 2 To last_row_ws8
'            On Error GoTo Error
'            ws8.Cells(i, 10).Formula = ws8.Cells(i, 2) & ws8.Cells(i, 3)
'            ws8.Cells(i, 11).Formula = "'" & .VLookup(ws8.Cells(i, 10), ws10.columns("C:G"), 2, False)
'            ws8.Cells(i, 12).Formula = .VLookup(ws8.Cells(i, 10), ws10.columns("C:G"), 3, False)
'            ws8.Cells(i, 13).Formula = "'" & .VLookup(ws8.Cells(i, 10), ws10.columns("C:G"), 4, False)
'            ws8.Cells(i, 14).Formula = .VLookup(ws8.Cells(i, 10), ws10.columns("C:G"), 5, False)
'            On Error GoTo 0
'
'            On Error GoTo syohinError
'            If Len(ws8.Cells(i, 5)) < 13 Then
'                ws8.Cells(i, 15).Formula = "'03"
'                ws8.Cells(i, 16).Formula = .VLookup(ws8.Cells(i, 5) * 1, ws4.columns("A:D"), 2, False)
'                ws8.Cells(i, 17).Formula = .VLookup(ws8.Cells(i, 5) * 1, ws4.columns("A:D"), 3, False)
'            Else
'                ws8.Cells(i, 15).Formula = "'02"
'                ws8.Cells(i, 16).Formula = .VLookup(ws8.Cells(i, 5) * 1, ws4.columns("A:D"), 2, False)
'            End If
'            On Error GoTo 0
'        Next i
'    End With
'
'    ws3.Rows("2:3000").Delete
'    ws8_LR_AC = ws8.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To ws8_LR_AC
'        If ws8.Cells(i, 11) = "" Then
'            MsgBox "返品用MSTを追加してください"
'            ws10.Activate
'            Exit Sub
'        ElseIf ws8.Cells(i, 16) = "" Then
'            MsgBox "商品MSTを追加してください"
'            ws4.Activate
'            Exit Sub
'        Else
'            ws3.Cells(i, 5).Value = ws8.Cells(i, 11)
'            ws3.Cells(i, 6).Value = ws8.Cells(i, 12)
'            ws3.Cells(i, 7).Value = ws8.Cells(i, 13)
'            ws3.Cells(i, 8).Value = ws8.Cells(i, 14)
'            If ws8.Cells(i, 15) = "02" Then
'                ws3.Cells(i, 9).Value = "'00000999"
'            Else
'                ws3.Cells(i, 9).Value = "'00000998"
'            End If
'            ws3.Cells(i, 10).Value = ws8.Cells(i, 1)
'            ws3.Cells(i, 11).Value = ws8.Cells(i, 15)
'            ws3.Cells(i, 12).Value = ws8.Cells(i, 16)
'            ws3.Cells(i, 14).Value = ws8.Cells(i, 17)
'            ws3.Cells(i, 18).Value = ws8.Cells(i, 7)
'            ws3.Cells(i, 19).Value = ws8.Cells(i, 8)
'            ws3.Cells(i, 20).Value = ws8.Cells(i, 9)
'            ws3.Cells(i, 25).Value = ws8.Cells(i, 18)
'
'        End If
'
'    Next i
'
'    If WorksheetFunction.Sum(ws8.columns("I")) <> WorksheetFunction.Sum(ws3.columns("T")) Then
'        MsgBox "返品金額計算合いません。"
'    End If
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
'    Dim SAVE_PATH As String
'    Dim AWB_NAME As String
'    AWB_NAME = Mid(ws3.Cells(2, 1), 3, 6) & "菊屋返品"
'
'    ' 保存先とファイル名を指定（例：Desktopに保存）
'    SAVE_PATH = Environ("USERPROFILE") & "\Desktop\" & AWB_NAME
'
'    ' ワークブックを保存
'    ActiveWorkbook.SaveAs Filename:=SAVE_PATH
'    Application.DisplayAlerts = True
'
'    MsgBox "既に計上済みの伝票を削除してください"
'    ' ワークブックを閉じる
'    'ActiveWorkbook.Close
'
'
'
'    Exit Sub
'
'Error:
'    Dim ws10_LR_AC As Long
'    ws10_LR_AC = ws10.Cells(Rows.Count, 1).End(xlUp).Row
'    ws10.Cells(ws10_LR_AC + 1, 1).Value = ws8.Cells(i, 2)
'    ws10.Cells(ws10_LR_AC + 1, 2).Value = ws8.Cells(i, 3)
'    ws10.Cells(ws10_LR_AC, 3).Copy ws10.Cells(ws10_LR_AC + 1, 3)
'    Resume Next
'
'syohinError:
'    Dim ws4_LR_AC As Long
'    ws4_LR_AC = ws4.Cells(Rows.Count, 1).End(xlUp).Row
'    ws4.Cells(ws4_LR_AC + 1, 1).Value = ws8.Cells(i, 5) * 1
'    Resume Next


'Sub 返品処理_1()
'
'    Dim WS As Worksheet
'
'    ' ワークシートをループして、オートフィルタが設定されているか確認する
'    For Each WS In ThisWorkbook.Worksheets
'        If WS.AutoFilterMode = True Then
'            WS.AutoFilterMode = False    ' オートフィルタを解除する
'        End If
'    Next WS
'
'    Dim WS3 As Worksheet
'    Set WS3 = ThisWorkbook.Worksheets("売上取込用")
'    Dim WS4 As Worksheet
'    Set WS4 = ThisWorkbook.Worksheets("商品MST")
'    Dim WS7 As Worksheet
'    Set WS7 = ThisWorkbook.Worksheets("支払明細書")
'    Dim WS8 As Worksheet
'    Set WS8 = ThisWorkbook.Worksheets("加工")
'    Dim WS10 As Worksheet
'    Set WS10 = ThisWorkbook.Worksheets("返品用MST")
'
'    On Error Resume Next    ' エラーが発生しても続行する
'    WS7.Cells.ClearFormats
'    On Error GoTo 0    ' エラーハンドリングを元に戻す
'
'    ' セルに対する個別のスタイルの削除
'    WS7.Cells.Style = "Normal"
'
'    Dim qt As QueryTable
'    Dim FileToOpen As Variant
'
'    WS7.Cells.ClearContents
'    ' CSVファイルを開く
'    FileToOpen = Application.GetOpenFilename("支払明細書CSVファイル (*.csv),*.csv", , "支払明細書CSVファイルを選択してください")
'    If FileToOpen <> "False" Then    ' ファイルが選択された場合のみ処理を行う
'        ' クエリテーブルを作成
'        Set qt = WS7.QueryTables.Add(Connection:="TEXT;" & FileToOpen, Destination:=WS7.Range("$A$1"))
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
'        For Each qt In WS7.QueryTables
'            qt.Delete
'        Next qt
'
'        ' 接続を削除
'        Dim cn As WorkbookConnection
'        For Each cn In ThisWorkbook.Connections
'            cn.Delete
'        Next cn
'    Else
'        Exit Sub
'    End If
'
'
'    WS8.Cells.ClearContents
'    WS7.Range("D1:L1").Copy WS8.Range("A1")
'
'    Dim WS7_LR_CC As Long
'    WS7_LR_CC = WS7.Cells(Rows.Count, 3).End(xlUp).Row
'    Dim WS8_LR_AC As Long
'    WS8_LR_AC = WS8.Cells(Rows.Count, 1).End(xlUp).Row
'
'
'    Dim i As Long
'    For i = 2 To WS7_LR_CC
'        If (WS7.Cells(i, 7) = "直仕返品") Or (WS7.Cells(i, 7) = "仕入返品") Or (WS7.Cells(i, 7) = "仕入返訂") Or (WS7.Cells(i, 7) = "直仕返訂") Then
'            'Debug.Print WS7.Cells(i, 4)
'            WS8.Cells(WS8_LR_AC + 1, 1).Value = WS7.Cells(i, 4)
'            WS8.Cells(WS8_LR_AC + 1, 2).Value = WS7.Cells(i, 5)
'            WS8.Cells(WS8_LR_AC + 1, 3).Value = WS7.Cells(i, 6)
'            WS8.Cells(WS8_LR_AC + 1, 4).Value = WS7.Cells(i, 7)
'            WS8.Cells(WS8_LR_AC + 1, 5).Value = WS7.Cells(i, 8)
'            WS8.Cells(WS8_LR_AC + 1, 6).Value = WS7.Cells(i, 9)
'            WS8.Cells(WS8_LR_AC + 1, 7).Value = WS7.Cells(i, 10)
'            WS8.Cells(WS8_LR_AC + 1, 8).Value = WS7.Cells(i, 11)
'            WS8.Cells(WS8_LR_AC + 1, 9).Value = WS7.Cells(i, 12)
'            WS8.Cells(WS8_LR_AC + 1, 18).Value = WS7.Cells(i, 16)
'
'            WS8_LR_AC = WS8.Cells(Rows.Count, 1).End(xlUp).Row
'        End If
'    Next i
'
'    'MsgBox "削除行を確認してください"
'    With WorksheetFunction
'
'        WS8_LR_AC = WS8.Cells(Rows.Count, 1).End(xlUp).Row
'        For i = 2 To WS8_LR_AC
'            On Error GoTo Error
'            WS8.Cells(i, 10).Formula = WS8.Cells(i, 2) & WS8.Cells(i, 3)
'            WS8.Cells(i, 11).Formula = "'" & .VLookup(WS8.Cells(i, 10), WS10.Columns("C:G"), 2, False)
'            WS8.Cells(i, 12).Formula = .VLookup(WS8.Cells(i, 10), WS10.Columns("C:G"), 3, False)
'            WS8.Cells(i, 13).Formula = "'" & .VLookup(WS8.Cells(i, 10), WS10.Columns("C:G"), 4, False)
'            WS8.Cells(i, 14).Formula = .VLookup(WS8.Cells(i, 10), WS10.Columns("C:G"), 5, False)
'            On Error GoTo 0
'
'            On Error GoTo syohinError
'            If Len(WS8.Cells(i, 5)) < 13 Then
'                WS8.Cells(i, 15).Formula = "'03"
'                WS8.Cells(i, 16).Formula = .VLookup(WS8.Cells(i, 5) * 1, WS4.Columns("A:D"), 2, False)
'                WS8.Cells(i, 17).Formula = .VLookup(WS8.Cells(i, 5) * 1, WS4.Columns("A:D"), 3, False)
'            Else
'                WS8.Cells(i, 15).Formula = "'02"
'                WS8.Cells(i, 16).Formula = .VLookup(WS8.Cells(i, 5) * 1, WS4.Columns("A:D"), 2, False)
'            End If
'            On Error GoTo 0
'        Next i
'    End With
'
'    WS3.Rows("2:3000").Delete
'    WS8_LR_AC = WS8.Cells(Rows.Count, 1).End(xlUp).Row
'    For i = 2 To WS8_LR_AC
'        If WS8.Cells(i, 11) = "" Then
'            MsgBox "返品用MSTを追加してください"
'            WS10.Activate
'            Exit Sub
'        ElseIf WS8.Cells(i, 16) = "" Then
'            MsgBox "商品MSTを追加してください"
'            WS4.Activate
'            Exit Sub
'        Else
'            WS3.Cells(i, 5).Value = WS8.Cells(i, 11)
'            WS3.Cells(i, 6).Value = WS8.Cells(i, 12)
'            WS3.Cells(i, 7).Value = WS8.Cells(i, 13)
'            WS3.Cells(i, 8).Value = WS8.Cells(i, 14)
'            If WS8.Cells(i, 15) = "02" Then
'                WS3.Cells(i, 9).Value = "'00000999"
'            Else
'                WS3.Cells(i, 9).Value = "'00000998"
'            End If
'            WS3.Cells(i, 10).Value = WS8.Cells(i, 1)
'            WS3.Cells(i, 11).Value = WS8.Cells(i, 15)
'            WS3.Cells(i, 12).Value = WS8.Cells(i, 16)
'            WS3.Cells(i, 14).Value = WS8.Cells(i, 17)
'            WS3.Cells(i, 18).Value = WS8.Cells(i, 7)
'            WS3.Cells(i, 19).Value = WS8.Cells(i, 8)
'            WS3.Cells(i, 20).Value = WS8.Cells(i, 9)
'            WS3.Cells(i, 25).Value = WS8.Cells(i, 18)
'
'        End If
'
'    Next i
'
'    If WorksheetFunction.Sum(WS8.Columns("I")) <> WorksheetFunction.Sum(WS3.Columns("T")) Then
'        MsgBox "返品金額計算合いません。"
'    End If
'
'    Dim IB As String
'    Dim WS3_LR_EC As Long
'    IB = InputBox("日付入力" & vbLf & vbLf & "20240101")
'
'    If IB = "" Then
'
'        Exit Sub
'    Else
'        WS3.Cells(2, 1).Value = IB
'        WS3_LR_EC = WS3.Cells(Rows.Count, 5).End(xlUp).Row
'        WS3.Cells(2, 1).Copy WS3.Range(WS3.Cells(2, 1), WS3.Cells(WS3_LR_EC, 2))
'        WS3.Cells(2, 1).Copy WS3.Range(WS3.Cells(2, 24), WS3.Cells(WS3_LR_EC, 24))
'    End If
'    WS3.Copy
'
'    Application.DisplayAlerts = False
'    Dim SAVE_PATH As String
'    Dim AWB_NAME As String
'    AWB_NAME = Mid(WS3.Cells(2, 1), 3, 6) & "菊屋返品"
'
'    ' 保存先とファイル名を指定（例：Desktopに保存）
'    SAVE_PATH = Environ("USERPROFILE") & "\Desktop\" & AWB_NAME
'
'    ' ワークブックを保存
'    ActiveWorkbook.SaveAs Filename:=SAVE_PATH
'    Application.DisplayAlerts = True
'
'    MsgBox "既に計上済みの伝票を削除してください"
'    ' ワークブックを閉じる
'    'ActiveWorkbook.Close
'
'
'
'    Exit Sub
'
'Error:
'    Dim WS10_LR_AC As Long
'    WS10_LR_AC = WS10.Cells(Rows.Count, 1).End(xlUp).Row
'    WS10.Cells(WS10_LR_AC + 1, 1).Value = WS8.Cells(i, 2)
'    WS10.Cells(WS10_LR_AC + 1, 2).Value = WS8.Cells(i, 3)
'    WS10.Cells(WS10_LR_AC, 3).Copy WS10.Cells(WS10_LR_AC + 1, 3)
'    Resume Next
'
'syohinError:
'    Dim WS4_LR_AC As Long
'    WS4_LR_AC = WS4.Cells(Rows.Count, 1).End(xlUp).Row
'    WS4.Cells(WS4_LR_AC + 1, 1).Value = WS8.Cells(i, 5) * 1
'    Resume Next
'End Sub