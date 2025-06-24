Option Explicit

Public Sub AO商品マスタ取込()

    Dim master_path As String
    master_path = "\\192.168.1.9\業務\　在庫表連携\商品マスタ_データ取込.xlsx"

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")

    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "manual")
    Call toggle_sheet_protection(ws1, "un_protect")

    Dim master_wb As Workbook
    Set master_wb = Workbooks.Open(master_path)


    Dim aws As Worksheet
    Set aws = master_wb.Worksheets("Sheet1")

    Dim data As Variant
    Dim aws_last_row As Long
    Dim aws_last_col As Long

    aws_last_row = aws.Cells(aws.Rows.Count, 1).End(xlUp).Row
    aws_last_col = aws.Cells(1, aws.Columns.Count).End(xlToLeft).Column

    data = aws.Range(aws.Cells(1, 1), aws.Cells(aws_last_row, aws_last_col)).Value

    Dim ws1_lastrow_d_col As Long
    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
    
    With ws1
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    ws1.AutoFilterMode = False

    ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_lastrow_d_col)).AutoFilter

    ws1.Rows("1:1").Hidden = True

    Dim i As Long
    Dim match_index As Variant
    Dim results() As Variant
    ReDim results(5 To ws1_lastrow_d_col, 1 To 6)    ' 2次元配列に変更

    Dim indexArray As Variant
    indexArray = Application.Index(data, 0, 1)

    For i = 5 To ws1_lastrow_d_col

        match_index = Application.Match(ws1.Cells(i, "D").Value, indexArray, 0)

 
    
        If Not IsError(match_index) And ws1.Cells(i, "D").Value <> "" Then
            results(i, 1) = data(match_index, 43)    'JAN
            results(i, 2) = data(match_index, 3) & data(match_index, 4)    '商品名
            results(i, 3) = data(match_index, 10)    '棚卸単価
            results(i, 4) = data(match_index, 11)    '上代
            results(i, 5) = data(match_index, 23)    '外箱入数
            results(i, 6) = data(match_index, 56)    '内箱入数
        ElseIf ws1.Cells(i, "D").Value = "" And ws1.Cells(i, "E").Value <> "ここからEC専用" Then

            Exit For
        ElseIf ws1.Cells(i, "E").Value = "" Then
       
            MsgBox ws1.Cells(i, "D").Value & "見当たりません"
            Exit For
        End If

    Next i

    For i = 5 To ws1_lastrow_d_col
        If Not IsEmpty(results(i, 1)) Then
            ws1.Cells(i, "C").Value = results(i, 1)    ' 単価情報をE列に貼り付け
        End If
        If Not IsEmpty(results(i, 2)) Then
            ws1.Cells(i, "E").Value = results(i, 2)    ' 商品名をF列に貼り付け
        End If
        If Not IsEmpty(results(i, 3)) Then
            ws1.Cells(i, "G").Value = results(i, 3)    ' 棚卸単価をG列に貼り付け
        End If
        If Not IsEmpty(results(i, 4)) Then
            ws1.Cells(i, "H").Value = results(i, 4)    ' 上代をH列に貼り付け
        End If
        If Not IsEmpty(results(i, 5)) Then
            ws1.Cells(i, "I").Value = results(i, 5)    ' 外箱入数をI列に貼り付け
        End If
        If Not IsEmpty(results(i, 6)) Then
            ws1.Cells(i, "J").Value = results(i, 6)    ' 内箱入数をJ列に貼り付け
        End If
    Next i

    master_wb.Close SaveChanges:=False

    Call toggle_screen_update(True)
    Call set_calculation_mode(ws1, "automatic")
    Call toggle_sheet_protection(ws1, "protect")

    Call 数式コピー

    MsgBox "更新完了"

End Sub

'With WorksheetFunction
'
'    AWBS_商品CD列_LR = AWBS.Cells(AWBS.Rows.Count, 商品CD列).End(xlUp).Row
'
'    For FOR1 = 2 To AWBS_商品CD列_LR
'        AWBS.Cells(FOR1, 1).Value = .Clean(AWBS.Cells(FOR1, 1))
'        AWBS.Cells(FOR1, 54).Value = .Clean(AWBS.Cells(FOR1, 54))
'    Next FOR1
'
'    S1.Unprotect
'
'    S1.Range(S1.Cells(データ開始行, 商品名列), S1.Cells(S1_商品CD列_LR - 1, 商品名列)).ClearContents
'    S1.Range(S1.Cells(データ開始行, 原価列), S1.Cells(S1_商品CD列_LR - 1, 原価列)).ClearContents
'    S1.Range(S1.Cells(データ開始行, 入数列), S1.Cells(S1_商品CD列_LR - 1, 入数列)).ClearContents
'    S1.Range(S1.Cells(データ開始行, JAN列), S1.Cells(S1_商品CD列_LR - 1, JAN列)).ClearContents
'    S1.Range(S1.Cells(データ開始行, 参考上代列), S1.Cells(S1_商品CD列_LR - 1, 参考上代列)).ClearContents
'
'
'    For FOR1 = データ開始行 To S1_商品CD列_LR - 1
'        On Error Resume Next
'        S1.Cells(FOR1, 商品名列).Value = _
          '                                     .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 2) & _
          '                                     .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 3)
'        S1.Cells(FOR1, 原価列).Value = _
          '                                    .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 16)
'        S1.Cells(FOR1, 入数列).Value = _
          '                                    .Clean(.Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 11)) * 1
'
'        If .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 14) = 0 _
          '           And .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 54) <> 0 Then
'
'            S1.Cells(FOR1, 参考上代列).Value = _
              '                                          "参 " & Format(.Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 54), "#,##0;[赤](#,##0)")
'        Else
'            S1.Cells(FOR1, 参考上代列).Value = _
              '                                          .Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 14)
'        End If
'
'        S1.Cells(FOR1, JAN列).Value = _
          '                                     .Clean(.Index(AWBS.Cells, .Match(S1.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 5)) * 1
'
'
'        On Error GoTo 0
'
'    Next FOR1
'End With
'
'For FOR1 = データ開始行 To S1_商品CD列_LR - 1
'    If S1.Cells(FOR1, 原価列) = "" Then
'        S1.Cells(FOR1, 原価列).Value = 0
'    End If
'    If S1.Cells(FOR1, 入数列) = "" Then
'        S1.Cells(FOR1, 入数列).Value = 1
'    End If
'Next FOR1
'
'S1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
  '         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
  '           AllowDeletingColumns:=True, AllowFiltering:=True
'
'S8.Unprotect
'
'With WorksheetFunction
'
'    S8.Range(S8.Cells(データ開始行, 商品名列), S8.Cells(S8_商品CD列_LR - 1, 商品名列)).ClearContents
'    S8.Range(S8.Cells(データ開始行, 原価列), S8.Cells(S8_商品CD列_LR - 1, 原価列)).ClearContents
'    S8.Range(S8.Cells(データ開始行, 入数列), S8.Cells(S8_商品CD列_LR - 1, 入数列)).ClearContents
'    S8.Range(S8.Cells(データ開始行, JAN列), S8.Cells(S8_商品CD列_LR - 1, JAN列)).ClearContents
'    S8.Range(S8.Cells(データ開始行, 参考上代列), S8.Cells(S8_商品CD列_LR - 1, 参考上代列)).ClearContents
'
'    For FOR1 = データ開始行 To S8_商品CD列_LR - 1
'        On Error Resume Next
'        S8.Cells(FOR1, 商品名列).Value = _
          '                                     .Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 2) & _
          '                                     .Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 3)
'
'        S8.Cells(FOR1, 原価列).Value = _
          '                                    .Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 16)
'        S8.Cells(FOR1, 入数列).Value = _
          '                                    .Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 11)
'        S8.Cells(FOR1, 参考上代列).Value = _
          '                                      .Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 14)
'        S8.Cells(FOR1, JAN列).Value = _
          '                                     .Clean(.Index(AWBS.Cells, .Match(S8.Cells(FOR1, 商品CD列), AWBS.Range("A:A"), 0), 5)) * 1
'
'
'        On Error GoTo 0
'
'    Next FOR1
'
'End With
'
'For FOR1 = データ開始行 To S8_商品CD列_LR - 1
'    If S8.Cells(FOR1, 原価列) = "" Then
'        S8.Cells(FOR1, 原価列).Value = 0
'    End If
'    If S8.Cells(FOR1, 入数列) = "" Then
'        S8.Cells(FOR1, 入数列).Value = 1
'    End If
'Next FOR1
'
'S8.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
  '         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
  '           AllowDeletingColumns:=True, AllowFiltering:=True
'
''Else
''Application.ScreenUpdating = True
''MsgBox "ファイルが存在しません"
''Exit Sub
''End If
'
'AWB.Close SaveChanges:=False
'
'Dim S1_実在庫C_4R As Long
'Dim S1_受注残C_4R As Long
'Dim S1_引当在庫_4R As Long
'Dim S1_次回入荷後引当C_4R As Long
'Dim S1_次々回入荷後引当C_4R As Long
'Dim S1_月間累計C_4R As Long
'Dim S1_翌月合計C_4R As Long
'Dim S1_入荷合計C_4R As Long
'Dim S1_出庫合計C_4R As Long
'Dim S1_本日出荷C_4R As Long
'Dim S2 As Worksheet
'Set S2 = ThisWorkbook.Worksheets("予約")
'Dim S2_商品CDC_4R As Long
'Dim S2_入荷後引当C_4R As Long
'Dim S2_合計C_4R As Long
'Dim S2_生産予定C_4R As Long
'Dim S2_次回予約合計C_4R As Long
'Dim S2_生産予定2C_4R As Long
'Dim S2_次々回予約合計C_4R As Long
'Dim S2_次々回入荷後引当C_4R As Long
'
'S1.Unprotect
'S2.Unprotect
'S8.Unprotect
'
'Dim S1_LC_4R As Long
'Dim S2_LC_4R As Long
'Dim S8_LC_4R As Long
'
'S1_LC_4R = S1.Cells(4, S1.Columns.Count).End(xlToLeft).Column
'S2_LC_4R = S2.Cells(4, S2.Columns.Count).End(xlToLeft).Column
'S8_LC_4R = S8.Cells(4, S8.Columns.Count).End(xlToLeft).Column
'
'If S1.AutoFilterMode = True Then
'    S1.AutoFilterMode = False
'    S1.Range(S1.Cells(4, 1), S1.Cells(4, S1_LC_4R)).AutoFilter
'End If
'
'If S2.AutoFilterMode = True Then
'    S2.AutoFilterMode = False
'    S2.Range(S2.Cells(4, 1), S2.Cells(4, S2_LC_4R)).AutoFilter
'End If
'
'If S8.AutoFilterMode = True Then
'    S8.AutoFilterMode = False
'    S8.Range(S8.Cells(4, 1), S8.Cells(4, S8_LC_4R)).AutoFilter
'End If
'
''コピー前にシートを全表示
'
'With S2
'    .Rows.Hidden = False
'    .Columns.Hidden = False
'End With
'
'With S8
'    .Rows.Hidden = False
'    .Columns.Hidden = False
'End With
'
'With S1
'    .Rows.Hidden = False
'    .Columns.Hidden = False
'End With
'
'S1_商品CD列_LR = S1.Cells(S1.Rows.Count, 商品CD列).End(xlUp).Row
'S8_商品CD列_LR = S8.Cells(S8.Rows.Count, 商品CD列).End(xlUp).Row
'
'With WorksheetFunction
'    S1_実在庫C_4R = .Match("実在庫", S1.Range("4:4"), 0)
'    S1_受注残C_4R = .Match("受注残", S1.Range("4:4"), 0)
'    S1_引当在庫_4R = .Match("引当在庫", S1.Range("4:4"), 0)
'    S1_次回入荷後引当C_4R = .Match("次回入荷後引当", S1.Range("4:4"), 0)
'    S1_次々回入荷後引当C_4R = .Match("次々回入荷後引当", S1.Range("4:4"), 0)
'    S1_月間累計C_4R = .Match("月間累計", S1.Range("4:4"), 0)
'    S1_翌月合計C_4R = .Match("翌月合計", S1.Range("4:4"), 0)
'    S1_入荷合計C_4R = .Match("入荷合計", S1.Range("4:4"), 0)
'    S1_出庫合計C_4R = .Match("出庫合計", S1.Range("4:4"), 0)
'    S1_本日出荷C_4R = .Match("本日出荷", S1.Range("4:4"), 0)
'    S2_商品CDC_4R = .Match("商品CD", S2.Range("4:4"), 0)
'    S2_入荷後引当C_4R = .Match("入荷後引当", S2.Range("4:4"), 0)
'    S2_合計C_4R = .Match("合計", S2.Range("4:4"), 0)
'    S2_生産予定C_4R = .Match("生産予定", S2.Range("4:4"), 0)
'    S2_次回予約合計C_4R = .Match("次回予約合計", S2.Range("4:4"), 0)
'    S2_生産予定2C_4R = .Match("生産予定2", S2.Range("4:4"), 0)
'    S2_次々回予約合計C_4R = .Match("次々回予約合計", S2.Range("4:4"), 0)
'    S2_次々回入荷後引当C_4R = .Match("次々回入荷後引当", S2.Range("4:4"), 0)
'
'    S1.Cells(データ開始行, S1_実在庫C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_実在庫C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_実在庫C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Range(S1.Cells(データ開始行, S1_受注残C_4R), S1.Cells(データ開始行, S1_引当在庫_4R)).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_受注残C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_引当在庫_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_次回入荷後引当C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_次回入荷後引当C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_次回入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_次々回入荷後引当C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_次々回入荷後引当C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_次々回入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_月間累計C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_月間累計C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_月間累計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_翌月合計C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_翌月合計C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_翌月合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_入荷合計C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_入荷合計C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_入荷合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_出庫合計C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_出庫合計C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_出庫合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S1.Cells(データ開始行, S1_本日出荷C_4R).Copy
'    S1.Range(S1.Cells(データ開始行 + 1, S1_本日出荷C_4R), S1.Cells(S1_商品CD列_LR - 1, S1_本日出荷C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S2.Range(S2.Cells(データ開始行, S2_商品CDC_4R), S2.Cells(データ開始行, S2_入荷後引当C_4R)).Copy
'    S2.Range(S2.Cells(データ開始行 + 1, S2_商品CDC_4R), S2.Cells(S1_商品CD列_LR - 1, S2_入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S2.Range(S2.Cells(データ開始行, S2_合計C_4R), S2.Cells(データ開始行, S2_生産予定C_4R)).Copy
'    S2.Range(S2.Cells(データ開始行 + 1, S2_合計C_4R), S2.Cells(S1_商品CD列_LR - 1, S2_生産予定C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S2.Range(S2.Cells(データ開始行, S2_次回予約合計C_4R), S2.Cells(データ開始行, S2_生産予定2C_4R)).Copy
'    S2.Range(S2.Cells(データ開始行 + 1, S2_次回予約合計C_4R), S2.Cells(S1_商品CD列_LR - 1, S2_生産予定2C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S2.Range(S2.Cells(データ開始行, S2_次々回予約合計C_4R), S2.Cells(データ開始行, S2_次々回入荷後引当C_4R)).Copy
'    S2.Range(S2.Cells(データ開始行 + 1, S2_次々回予約合計C_4R), S2.Cells(S1_商品CD列_LR - 1, S2_次々回入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'End With
'
'Dim S8_実在庫C_4R As Long
'Dim S8_受注残C_4R As Long
'Dim S8_引当在庫_4R As Long
'Dim S8_次回入荷後引当C_4R As Long
'Dim S8_次々回入荷後引当C_4R As Long
'Dim S8_月間累計C_4R As Long
'Dim S8_翌月合計C_4R As Long
'Dim S8_入荷合計C_4R As Long
'Dim S8_出庫合計C_4R As Long
'Dim S8_本日出荷C_4R As Long
'
'With WorksheetFunction
'    S8_実在庫C_4R = .Match("実在庫", S8.Range("4:4"), 0)
'    S8_受注残C_4R = .Match("受注残", S8.Range("4:4"), 0)
'    S8_引当在庫_4R = .Match("引当在庫", S8.Range("4:4"), 0)
'    S8_次回入荷後引当C_4R = .Match("次回入荷後引当", S8.Range("4:4"), 0)
'    S8_次々回入荷後引当C_4R = .Match("次々回入荷後引当", S8.Range("4:4"), 0)
'    S8_月間累計C_4R = .Match("月間累計", S8.Range("4:4"), 0)
'    S8_翌月合計C_4R = .Match("翌月合計", S8.Range("4:4"), 0)
'    S8_入荷合計C_4R = .Match("入荷合計", S8.Range("4:4"), 0)
'    S8_出庫合計C_4R = .Match("出庫合計", S8.Range("4:4"), 0)
'    S8_本日出荷C_4R = .Match("本日出荷", S8.Range("4:4"), 0)
'
'    S8.Cells(データ開始行, S8_実在庫C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_実在庫C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_実在庫C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Range(S8.Cells(データ開始行, S8_受注残C_4R), S8.Cells(データ開始行, S8_引当在庫_4R)).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_受注残C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_引当在庫_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_次回入荷後引当C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_次回入荷後引当C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_次回入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_次々回入荷後引当C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_次々回入荷後引当C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_次々回入荷後引当C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_月間累計C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_月間累計C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_月間累計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_翌月合計C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_翌月合計C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_翌月合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_入荷合計C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_入荷合計C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_入荷合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_出庫合計C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_出庫合計C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_出庫合計C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'    S8.Cells(データ開始行, S8_本日出荷C_4R).Copy
'    S8.Range(S8.Cells(データ開始行 + 1, S8_本日出荷C_4R), S8.Cells(S8_商品CD列_LR - 1, S8_本日出荷C_4R)).PasteSpecial _
      '            Paste:=xlPasteFormulas
'
'End With
'
'S1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
  '         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
  '           AllowDeletingColumns:=True, AllowFiltering:=True
'
'S2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
  '         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
  '           AllowDeletingColumns:=True, AllowFiltering:=True
'
'S8.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
  '         , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
  '           AllowDeletingColumns:=True, AllowFiltering:=True
'
'S1.Activate
'Unload UserForm2
'Application.ScreenUpdating = True
'
'MsgBox "取込完了"
'



'End Sub