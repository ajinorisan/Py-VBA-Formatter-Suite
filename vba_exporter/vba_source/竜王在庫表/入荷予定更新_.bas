Option Explicit

Public Sub 入荷予定更新()

    Dim inport_path As String
    inport_path = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "輸入計画データを選択してください")

    If inport_path = "False" Then
        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbExclamation
        Exit Sub
    End If

    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("在庫表")

    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "manual")
    Call toggle_sheet_protection(ws1, "un_protect")


    Workbooks.Open Filename:= _
                   inport_path, ReadOnly:=True

    Dim inport_file As Workbook
    Set inport_file = ActiveWorkbook


    Dim aws As Worksheet
    Set aws = inport_file.Worksheets("在庫表連携")

    Dim aws2 As Worksheet

    Dim new_sheet_name As String
    new_sheet_name = "更新履歴"

    Dim true_false As Boolean
    true_false = False

    Dim ws As Worksheet
    For Each ws In inport_file.Worksheets
        If ws.Name = new_sheet_name Then
            Set aws2 = ws
            true_false = True
            Exit For
        End If
    Next ws

    If Not true_false Then
        Set aws2 = inport_file.Worksheets.Add
        aws2.Name = new_sheet_name
    End If

    Dim ws1_lastrow_d_col As Long
    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

    Dim ws1_4_row_lastcol As Long
    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

    If ws1.AutoFilterMode = True Then
        ws1.AutoFilterMode = False
        ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_4_row_lastcol)).AutoFilter
    End If

    With ws1
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    Dim i As Long

    Dim copy_data As Variant

    aws2.Cells.Delete
    aws2.Cells(1, 1) = WorksheetFunction.text(Date, "mm/dd")

    ReDim copy_data(1 To ws1_lastrow_d_col - 4, 1 To 6)
    Dim copy_index As Long
    copy_index = 1

    For i = 4 To ws1_lastrow_d_col - 1
        If InStr(ws1.Cells(i, "F").Value, "国内") = 0 Then

            copy_data(copy_index, 1) = ws1.Cells(i, "D").Value
            copy_data(copy_index, 2) = ws1.Cells(i, "E").Value
            copy_data(copy_index, 3) = ws1.Cells(i, "V").Value
            copy_data(copy_index, 4) = ws1.Cells(i, "W").Value
            copy_data(copy_index, 5) = ws1.Cells(i, "Y").Value
            copy_data(copy_index, 6) = ws1.Cells(i, "Z").Value

            copy_index = copy_index + 1

        End If
    Next i

    aws2.Range("A2").Resize(UBound(copy_data, 1), UBound(copy_data, 2)).Value = copy_data

    Application.CutCopyMode = False

    Dim aws2_last_col As Long
    aws2_last_col = aws2.Cells(2, aws2.Columns.Count).End(xlToLeft).Column

    Dim match_index1 As Variant
    Dim match_index2 As Variant

    Dim aws_Data As Variant
    Dim aws2_Data As Variant

    Dim aws_last_row_b_col As Long
    aws_last_row_b_col = aws.Cells(aws.Rows.Count, "B").End(xlUp).Row

    Dim aws2_last_row_a_col As Long
    aws2_last_row_a_col = aws2.Cells(aws2.Rows.Count, "A").End(xlUp).Row

    aws_Data = aws.Range(aws.Cells(1, "A"), aws.Cells(aws_last_row_b_col, "L")).Value
    aws2_Data = aws2.Range(aws2.Cells(1, "A"), aws2.Cells(aws2_last_row_a_col, "A")).Value

    For i = 3 To aws2_last_row_a_col
        match_index1 = Application.Match(aws2_Data(i, 1), Application.Index(aws_Data, 0, 2), 0)
        match_index2 = Application.Match(aws2_Data(i, 1), Application.Index(aws_Data, 0, 5), 0)

        If Not IsError(match_index1) Then
            aws2.Cells(i, "G").Value = aws_Data(match_index1, 1)
            aws2.Cells(i, "H").Value = aws_Data(match_index1, 3)
        End If

        If Not IsError(match_index2) Then
            aws2.Cells(i, "I").Value = aws_Data(match_index2, 4)
            aws2.Cells(i, "J").Value = aws_Data(match_index2, 6)
        End If

        If aws2.Cells(i, "C").Value <> aws2.Cells(i, "G").Value Or _
           aws2.Cells(i, "D").Value <> aws2.Cells(i, "H").Value Or _
           aws2.Cells(i, "E").Value <> aws2.Cells(i, "I").Value Or _
           aws2.Cells(i, "F").Value <> aws2.Cells(i, "J").Value Then
            aws2.Cells(i, "K").Value = "変更有"
        End If

        If aws2.Cells(i, "K").Value = "" Then
            aws2.Rows(i).Hidden = True
        End If
    Next i



    aws2.Cells(2, "G").formula = "変更予定日"
    aws2.Cells(2, "H").formula = "変更数量"
    aws2.Cells(2, "I").formula = "変更予定日2"
    aws2.Cells(2, "J").formula = "変更数量2"
    aws2.Columns("G").NumberFormat = "mm/dd"
    aws2.Columns("I").NumberFormat = "mm/dd"

    aws2.Columns("A:A").ColumnWidth = 17.86
    aws2.Columns("B:B").ColumnWidth = 20.71
    aws2.Range("C:C,E:E,G:G,I:I").ColumnWidth = 12.14
    aws2.Range("D:D,F:F,H:H,J:J").ColumnWidth = 10.71
    aws2.Cells.ShrinkToFit = True

    Dim target_range As Range
    Set target_range = aws2.Range(aws2.Cells(2, 1), aws2.Cells(aws2_last_row_a_col, "K"))

    With target_range.Borders
        .LineStyle = xlContinuous    ' 線のスタイルを設定
        .Weight = xlThin    ' 線の太さを設定
        .ColorIndex = xlAutomatic    ' 線の色を自動に設定
    End With
    aws2.Columns("K").Hidden = True

    aws2.Cells.Font.Name = "メイリオ"
    aws2.Cells.Font.Size = 8

    Application.DisplayAlerts = False

    If inport_path <> "False" Then

        Dim file_name As String
        file_name = Dir(inport_path)

        Dim output_path As String
        output_path = Left(inport_path, Len(inport_path) - Len(file_name)) & _
                      Format(Date, "YYYYMMDD") & "輸入計画更新.pdf"

        With aws2.PageSetup
            .Orientation = xlLandscape
            .FitToPagesWide = 1    ' 幅を1ページに設定
            .FitToPagesTall = False    ' 高さは自動調整
            .LeftMargin = Application.InchesToPoints(0.5)    ' 左余白を0.5インチに設定
            .RightMargin = Application.InchesToPoints(0.5)    ' 右余白を0.5インチに設定
            .TopMargin = Application.InchesToPoints(0.5)    ' 上余白を0.5インチに設定
            .BottomMargin = Application.InchesToPoints(0.5)    ' 下余白を0.5インチに設定
        End With

        aws2.ExportAsFixedFormat Type:=xlTypePDF, Filename:=output_path, Quality:= _
                                 xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                                 OpenAfterPublish:=True

    End If

    Dim match_index_D As Variant
    Dim match_index_E As Variant

    For i = 5 To ws1_lastrow_d_col - 1
        If InStr(ws1.Cells(i, "F"), "国内") = 0 Then
            ws1.Range(ws1.Cells(i, "V"), ws1.Cells(i, "W")).ClearContents
            ws1.Range(ws1.Cells(i, "Y"), ws1.Cells(i, "Z")).ClearContents
        End If

        match_index_D = Application.Match(ws1.Cells(i, "D").Value, aws.Range("B:B"), 0)
        match_index_E = Application.Match(ws1.Cells(i, "D").Value, aws.Range("E:E"), 0)


        If Not IsError(match_index_D) Then
            ws1.Cells(i, "V").Value = aws.Cells(match_index_D, 1).Value
            ws1.Cells(i, "W").Value = aws.Cells(match_index_D, 3).Value
        End If

        If Not IsError(match_index_E) Then
            ws1.Cells(i, "Y").Value = aws.Cells(match_index_E, 4).Value
            ws1.Cells(i, "Z").Value = aws.Cells(match_index_E, 6).Value
        End If

    Next i

    inport_file.Close
    Application.DisplayAlerts = True

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws1, "automatic")
    Call toggle_sheet_protection(ws1, "protect")


End Sub

'Sub 入荷予定更新()
'
''https://docs.google.com/spreadsheets/d/11V622sThDYEHTPk5GMnkMkzD77aealpXIGOYzjTW6Tc/export?format=xlsx
'
'    Dim 輸入計画パス As String
'    ' ファイル選択ダイアログを表示して、ユーザーにファイルを選ばせる
'    輸入計画パス = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "輸入計画データを選択してください")
'
'
'
'    ' ユーザーがキャンセルした場合の処理
'    If 輸入計画パス = "False" Then
'        MsgBox "ファイルが選択されませんでした。処理を中止します。", vbExclamation
'        Exit Sub
'    End If
'    Dim 商品CD列 As Long
'    商品CD列 = 4
'    Dim データ開始行 As Long
'    データ開始行 = 5
'    Dim 入荷予定日 As Long
'    入荷予定日 = 18
'    Dim 生産予定 As Long
'    生産予定 = 19
'    Dim 入荷予定日2 As Long
'    入荷予定日2 = 21
'    Dim 生産予定2 As Long
'    生産予定2 = 22
'    Dim 備考列 As Long
'    備考列 = 6
'
'    Application.ScreenUpdating = False
'    Application.Calculation = xlManual
'
'    Dim S1 As Worksheet
'    Set S1 = ThisWorkbook.Worksheets("在庫数")
'
'    S1.Select
'
'    Workbooks.Open Filename:= _
      '                   輸入計画パス, ReadOnly:=True
'
'    Dim AWB As Workbook
'    Set AWB = ActiveWorkbook
'    Dim AWBS As Worksheet
'    Set AWBS = ActiveWorkbook.Worksheets("在庫表連携")
'    Dim AWBS2 As Worksheet
'    Dim シート名 As String
'    シート名 = "更新履歴"
'
'    ' シートが存在するかどうかを確認するための変数
'    Dim シート存在 As Boolean
'    シート存在 = False
'
'    ' すべてのシートをループして、指定した名前のシートが存在するか確認
'    Dim ws As Worksheet
'    For Each ws In ActiveWorkbook.Worksheets
'        If ws.Name = シート名 Then
'            Set AWBS2 = ws
'            シート存在 = True
'            Exit For
'        End If
'    Next ws
'
'    ' 指定した名前のシートが存在しない場合、新しいシートを追加
'    If Not シート存在 Then
'        Set AWBS2 = ActiveWorkbook.Worksheets.Add
'        AWBS2.Name = シート名
'    End If
'
'
'
'    Dim i As Long
'
'    Dim S1_last_row As Long
'    S1_last_row = S1.Cells(S1.Rows.Count, 商品CD列).End(xlUp).Row
'
'    Dim S1_last_column As Long
'    S1_last_column = S1.Cells(データ開始行 - 1, S1.Columns.Count).End(xlToLeft).Column
'
'    S1.Unprotect
'
'    If S1.AutoFilterMode = True Then
'        S1.AutoFilterMode = False
'        S1.Range(S1.Cells(4, 1), S1.Cells(4, S1_last_column)).AutoFilter
'    End If
'
'    With S1
'        .Rows.Hidden = False
'        .Columns.Hidden = False
'    End With
'
'    AWBS2.Cells.Delete
'    AWBS2.Cells(1, 1) = WorksheetFunction.text(Date, "mm/dd")
'    For i = データ開始行 - 1 To S1_last_row - 1
'        If InStr(S1.Cells(i, 備考列), "国内") = 0 Then
'
'
'            S1.Range(S1.Cells(i, 4), S1.Cells(i, 5)).Copy _
              '                    AWBS2.Range(AWBS2.Cells(i - 2, 1), AWBS2.Cells(i - 2, 2))
'            S1.Range(S1.Cells(i, 入荷予定日), S1.Cells(i, 生産予定)).Copy _
              '                    AWBS2.Range(AWBS2.Cells(i - 2, 3), AWBS2.Cells(i - 2, 4))
'            S1.Range(S1.Cells(i, 入荷予定日2), S1.Cells(i, 生産予定2)).Copy _
              '                    AWBS2.Range(AWBS2.Cells(i - 2, 5), AWBS2.Cells(i - 2, 6))
'
'        End If
'    Next i
'
'    Dim AWBS2_LC As Long
'    AWBS2_LC = AWBS2.Cells(2, AWBS2.Columns.Count).End(xlToLeft).Column
'
'    For i = 2 To S1_last_row - 3
'
'        With WorksheetFunction
'            On Error Resume Next
'            AWBS2.Cells(i, 7).Value = .Index(AWBS.Cells, .Match(AWBS2.Cells(i, 1), AWBS.Range("B:B"), 0), 1)
'            AWBS2.Cells(i, 8).Value = .Index(AWBS.Cells, .Match(AWBS2.Cells(i, 1), AWBS.Range("B:B"), 0), 3)
'            AWBS2.Cells(i, 9).Value = .Index(AWBS.Cells, .Match(AWBS2.Cells(i, 1), AWBS.Range("E:E"), 0), 4)
'            AWBS2.Cells(i, 10).Value = .Index(AWBS.Cells, .Match(AWBS2.Cells(i, 1), AWBS.Range("E:E"), 0), 6)
'            On Error GoTo 0
'        End With
'        If AWBS2.Cells(i, 3) <> AWBS2.Cells(i, 7) Then
'            AWBS2.Cells(i, 11).Value = 1
'        End If
'        If AWBS2.Cells(i, 4) <> AWBS2.Cells(i, 8) Then
'            AWBS2.Cells(i, 11).Value = 1
'        End If
'        If AWBS2.Cells(i, 5) <> AWBS2.Cells(i, 9) Then
'            AWBS2.Cells(i, 11).Value = 1
'        End If
'        If AWBS2.Cells(i, 6) <> AWBS2.Cells(i, 10) Then
'            AWBS2.Cells(i, 11).Value = 1
'        End If
'        If AWBS2.Cells(i, 11) = "" Then
'            AWBS2.Rows(i).Hidden = True
'        End If
'    Next i
'
'    AWBS2.Cells(2, 7).formula = "変更竜王予定日"
'    AWBS2.Cells(2, 8).formula = "変更数量"
'    AWBS2.Cells(2, 9).formula = "変更竜王予定日2"
'    AWBS2.Cells(2, 10).formula = "変更数量2"
'    AWBS2.Columns(7).NumberFormat = "mm/dd"
'    AWBS2.Columns(9).NumberFormat = "mm/dd"
'    AWBS2.Columns("G:J").AutoFit
'
'    AWBS2.Cells.Font.Name = "メイリオ"
'    AWBS2.Cells.Font.Size = 8
'
'    Dim 出力パス As String
'    ' ユーザーがファイルを選択した場合
'    If 輸入計画パス <> "False" Then
'        ' 選択したファイルのパスからファイル名を取得
'        Dim ファイル名 As String
'        ファイル名 = Dir(輸入計画パス)
'
'        ' 出力パスを設定
'        出力パス = Left(輸入計画パス, Len(輸入計画パス) - Len(ファイル名)) & _
          '               Format(Date, "YYYYMMDD") & "輸入計画更新.pdf"
'
'        With AWBS2.PageSetup
'            .Orientation = xlLandscape
'        End With
'
'        Application.DisplayAlerts = False
'
'        ' PDFとしてエクスポート
'        AWBS2.ExportAsFixedFormat Type:=xlTypePDF, Filename:=出力パス, Quality:= _
          '                                  xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
          '                                  OpenAfterPublish:=True
'        Application.DisplayAlerts = True
'
'    End If
'
'    For i = データ開始行 To S1_last_row - 1
'        If InStr(S1.Cells(i, 備考列), "国内") = 0 Then
'            S1.Range(S1.Cells(i, 入荷予定日), S1.Cells(i, 生産予定)).ClearContents
'            S1.Range(S1.Cells(i, 入荷予定日2), S1.Cells(i, 生産予定2)).ClearContents
'        End If
'        With WorksheetFunction
'            On Error Resume Next
'            S1.Cells(i, 入荷予定日).Value = .Index(AWBS.Cells, .Match(S1.Cells(i, 4), AWBS.Range("B:B"), 0), 1)
'            S1.Cells(i, 生産予定).Value = .Index(AWBS.Cells, .Match(S1.Cells(i, 4), AWBS.Range("B:B"), 0), 3)
'            S1.Cells(i, 入荷予定日2).Value = .Index(AWBS.Cells, .Match(S1.Cells(i, 4), AWBS.Range("E:E"), 0), 4)
'            S1.Cells(i, 生産予定2).Value = .Index(AWBS.Cells, .Match(S1.Cells(i, 4), AWBS.Range("E:E"), 0), 6)
'            On Error GoTo 0
'        End With
'    Next i
'
'
'    S1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '               , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '               AllowDeletingColumns:=True, AllowFiltering:=True
'
'    Application.DisplayAlerts = False
'    AWB.Close
'    Application.DisplayAlerts = True
'
'    S1.Activate
'
'    Application.Calculation = xlAutomatic
'    Application.ScreenUpdating = True
'
'    MsgBox "輸入計画更新完了" & vbCrLf & "国内品、OEM品は手作業で入力してください"
'
'End Sub
