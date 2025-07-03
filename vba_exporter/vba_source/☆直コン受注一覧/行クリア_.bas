Option Explicit

Sub 行クリア()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    Application.ScreenUpdating = False

    With WS1  '①

        .Unprotect    'プロテクト解除
        On Error Resume Next
        .Cells.EntireColumn.Hidden = False    '列の非表示解除
        On Error GoTo 0
        'If .AutoFilterMode = True Then 'オートフィルタモードなら解除
        '.AutoFilterMode = False
        'End If
        Dim YESNO As Long

        YESNO = MsgBox("処理を行いますか？", vbYesNo + vbQuestion, "確認")

        If YESNO = vbYes Then
            WS1.Range(WS1.Cells(Selection(1).Row, 1), WS1.Cells(Selection(Selection.Count).Row, 8)).ClearContents
            WS1.Range(WS1.Cells(Selection(1).Row, 13), WS1.Cells(Selection(Selection.Count).Row, 17)).ClearContents
            WS1.Range(WS1.Cells(Selection(1).Row, 20), WS1.Cells(Selection(Selection.Count).Row, 24)).ClearContents
            WS1.Range(WS1.Cells(Selection(1).Row, 27), WS1.Cells(Selection(Selection.Count).Row, 30)).ClearContents
            WS1.Range(WS1.Cells(Selection(1).Row, 33), WS1.Cells(Selection(Selection.Count).Row, 45)).ClearContents

            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
            , AllowFiltering:=True    'プロテクト
            .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能
            Application.ScreenUpdating = True
            AppActivate Application.Caption
            Unload UserForm1
            '        MsgBox "処理終了"
        Else
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
            , AllowFiltering:=True    'プロテクト
            .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能
            Application.ScreenUpdating = True
            AppActivate Application.Caption
            Unload UserForm1
            '        MsgBox "処理を中断します"
            Exit Sub
        End If

    End With
End Sub