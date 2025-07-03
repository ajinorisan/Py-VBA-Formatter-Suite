Option Explicit

Public Sub 在庫共有更新()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")
    Set ws3 = ThisWorkbook.Worksheets("受注残")

    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "manual")
    Call toggle_sheet_protection(ws1, "un_protect")
    Call set_calculation_mode(ws2, "manual")
    Call toggle_sheet_protection(ws2, "un_protect")

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Dim ws1_4_row_last_col As Long
    ws1_4_row_last_col = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

    Dim ws2_4_row_last_col As Long
    ws2_4_row_last_col = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

    With ws1

        If .AutoFilterMode = True Then
            .AutoFilterMode = False
            .Range(.Cells(4, 1), .Cells(4, ws1_4_row_last_col)).AutoFilter
        End If

        .Rows.Hidden = False
        .Columns.Hidden = False

        .Range("E3").Value = "更新" & Now()
    End With

    With ws2

        If .AutoFilterMode = True Then
            .AutoFilterMode = False
            .Range(.Cells(4, 1), .Cells(4, ws2_4_row_last_col)).AutoFilter
        End If

        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    With ws3

        If .AutoFilterMode = True Then
            .AutoFilterMode = False
        End If

        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    Worksheets(Array("在庫表", "予約", "受注残")).Copy

    Dim aws1 As Worksheet
    Set aws1 = ActiveWorkbook.Worksheets("在庫表")
    Dim aws2 As Worksheet
    Set aws2 = ActiveWorkbook.Worksheets("予約")

    Dim gessyo As Long
    gessyo = Application.Match("月初", ws1.Rows(4), 0)

    Dim honjitu_syukka As Long
    honjitu_syukka = Application.Match("本日出荷", ws1.Rows(4), 0)

    aws1.Range(aws1.Columns(gessyo), aws1.Columns(gessyo + 1)).Copy
    aws1.Range(aws1.Columns(gessyo), aws1.Columns(gessyo + 1)).PasteSpecial Paste:=xlPasteValues

    aws1.Columns(honjitu_syukka).Copy
    aws1.Columns(honjitu_syukka).PasteSpecial Paste:=xlPasteValues

    aws1.Cells(2, 1).ClearContents
    aws1.Rows(1).ClearContents

    Dim obj As Object

    For Each obj In aws1.OLEObjects
        obj.Delete
    Next obj
    For Each obj In aws1.Shapes
        If obj.Type = msoFormControl Then
            obj.Delete
        End If
    Next obj

    For Each obj In aws2.OLEObjects
        obj.Delete
    Next obj
    For Each obj In aws2.Shapes
        If obj.Type = msoFormControl Then
            obj.Delete
        End If
    Next obj

    Dim copy_path As String
    copy_path = "\\192.168.1.218\業務\在庫表コピー本社用"
    ChDir copy_path
    ActiveWorkbook.SaveAs Filename:= _
                          copy_path & "\竜王在庫表.xlsx", FileFormat:= _
                          xlOpenXMLWorkbook, WriteResPassword:="0413", CreateBackup:=False


    ActiveWorkbook.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call toggle_screen_update(False)
    Call set_calculation_mode(ws1, "automatic")
    Call toggle_sheet_protection(ws1, "protect")
    Call set_calculation_mode(ws2, "automatic")
    Call toggle_sheet_protection(ws2, "protect")


End Sub

'Sub 本社用在庫表更新()
'
'Dim S1 As Worksheet
'Dim S2 As Worksheet
'Dim S3 As Worksheet
'Dim S8 As Worksheet
'
'
'Set S1 = ThisWorkbook.Worksheets("在庫数")
'Set S2 = ThisWorkbook.Worksheets("予約")
'Set S3 = ThisWorkbook.Worksheets("受注残")
'Set S8 = ThisWorkbook.Worksheets("OEM在庫")
'
'
'Dim S1_本日出荷C_4R As Long
'Dim S8_本日出荷C_4R As Long
'Dim S1_AO在庫C_4R As Long
'Dim S8_AO在庫C_4R As Long
'Dim S1_LC_4R As Long
'Dim S2_LC_4R As Long
'Dim S3_LC_4R As Long
'Dim S8_LC_4R As Long
'
'S1_LC_4R = S1.Cells(4, S1.Columns.Count).End(xlToLeft).Column
'S2_LC_4R = S2.Cells(4, S2.Columns.Count).End(xlToLeft).Column
'S3_LC_4R = S3.Cells(4, S3.Columns.Count).End(xlToLeft).Column
'S8_LC_4R = S8.Cells(4, S8.Columns.Count).End(xlToLeft).Column
'
'S1_AO在庫C_4R = 11
' S8_AO在庫C_4R = 11
'
''保存先のフルパス
'Dim 竜王WS在庫表保存先パス As String
''竜王WS在庫表保存先パス = "\\192.168.1.9\業務\竜王在庫"
'竜王WS在庫表保存先パス = "\\192.168.1.218\業務\在庫表コピー本社用"
'
''画面変更を止める
'Application.ScreenUpdating = False
'
'S1.Unprotect
'S2.Unprotect
'S3.Unprotect
'S8.Unprotect
'
'   If S1.AutoFilterMode = True Then
'S1.AutoFilterMode = False
'S1.Range(S1.Cells(4, 1), S1.Cells(4, S1_LC_4R)).AutoFilter
'End If
'
' If S2.AutoFilterMode = True Then
'S2.AutoFilterMode = False
'S2.Range(S2.Cells(4, 1), S2.Cells(4, S2_LC_4R)).AutoFilter
'End If
'
' If S3.AutoFilterMode = True Then
'S3.AutoFilterMode = False
'S3.Range(S3.Cells(4, 1), S3.Cells(4, S3_LC_4R)).AutoFilter
'End If
'
'If S8.AutoFilterMode = True Then
'S8.AutoFilterMode = False
'S8.Range(S8.Cells(4, 1), S8.Cells(4, S8_LC_4R)).AutoFilter
'End If
'
''コピー前にシートを全表示
'With S3
'.Rows.Hidden = False
'    .Columns.Hidden = False
'
'    End With
'
'With S2
'.Rows.Hidden = False
'    .Columns.Hidden = False
'    End With
'
'    With S8
'.Rows.Hidden = False
'    .Columns.Hidden = False
'    End With
'
'    With S1
'    .Rows.Hidden = False
'    .Columns.Hidden = False
'
'    '更新時刻を入力
'       .Range("E3").Value = "更新" & Now()
''       .Range("E3").Value = Now()
'       End With
'
'       S1_LC_4R = S1.Cells(4, S1.Columns.Count).End(xlToLeft).Column
'       S8_LC_4R = S8.Cells(4, S8.Columns.Count).End(xlToLeft).Column
'
'       With WorksheetFunction
'S1_本日出荷C_4R = .Match("本日出荷", S1.Range(S1.Cells(4, 1), S1.Cells(4, S1_LC_4R)), 0)
'S8_本日出荷C_4R = .Match("本日出荷", S8.Range(S8.Cells(4, 1), S8.Cells(4, S8_LC_4R)), 0)
'
'End With
'    'シートをコピー
'    Worksheets(Array("在庫数", "予約", "受注残", "OEM在庫")).Copy
'
'    Dim C_S1 As Worksheet
'    Set C_S1 = ActiveWorkbook.Worksheets("在庫数")
'    Dim C_S8 As Worksheet
'    Set C_S8 = ActiveWorkbook.Worksheets("OEM在庫")
'
'    Application.EnableEvents = False
'    C_S1.Columns(S1_本日出荷C_4R).Copy
'    C_S1.Columns(S1_本日出荷C_4R).PasteSpecial Paste:=xlPasteValues
'
'    C_S8.Columns(S8_本日出荷C_4R).Copy
'    C_S8.Columns(S8_本日出荷C_4R).PasteSpecial Paste:=xlPasteValues
'
'    C_S1.Columns(S1_AO在庫C_4R).Copy
'    C_S1.Columns(S8_AO在庫C_4R).PasteSpecial Paste:=xlPasteValues
'
'    C_S8.Columns(S1_AO在庫C_4R).Copy
'    C_S8.Columns(S8_AO在庫C_4R).PasteSpecial Paste:=xlPasteValues
'
'    Application.EnableEvents = True
'
''   C_S1.Shapes.Range(Array("CommandButton23")).Delete
''   C_S1.Shapes.Range(Array("CommandButton1")).Delete
''   C_S1.Shapes.Range(Array("CommandButton23")).Delete
'
'    '保存先指定
'    ChDir 竜王WS在庫表保存先パス
'    '上書きするかの警告を非表示
'    Application.DisplayAlerts = False
'    On Error Resume Next
'    'コピーを書込みパスワード付で保存
'  ActiveWorkbook.SaveAs Filename:= _
    '                          竜王WS在庫表保存先パス & "\竜王在庫表.xlsx", FileFormat:= _
    '                          xlOpenXMLWorkbook, WriteResPassword:="0413", CreateBackup:=False
'        On Error GoTo 0
'        'コピーを閉じる
'        ActiveWindow.Close
'
'        '更新時刻をクリア
''        S1.Range("E3").ClearContents
'
'    S1.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
      '        , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
      '        AllowDeletingColumns:=True, AllowFiltering:=True
'
'        S2.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
          '        , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
          '        AllowDeletingColumns:=True, AllowFiltering:=True
'
'        S8.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
          '        , AllowFormattingCells:=True, AllowInsertingColumns:=True, _
          '        AllowDeletingColumns:=True, AllowFiltering:=True
'
'        '警告を再表示
'        Application.DisplayAlerts = True
'        '画面変更を開始
'                Application.ScreenUpdating = True
'                MsgBox "更新終了"
'
'End Sub
