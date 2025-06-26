Option Explicit

Sub 分納()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    If Selection.Rows.Count <> 1 Then
        MsgBox "1行ごとに処理してください"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    With WS1  '①

        .Unprotect    'プロテクト解除
        On Error Resume Next
        .Cells.EntireColumn.Hidden = False    '列の非表示解除
        On Error GoTo 0

        Dim YESNO As Long






        YESNO = MsgBox("分納処理を行いますか？", vbYesNo + vbQuestion, "確認")



        If YESNO = vbYes Then

            Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
            S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

            If .AutoFilterMode = True Then    'オートフィルタモードなら解除
                .AutoFilterMode = False
            End If

            Dim S1_A_LASTROW As Long

            S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row

            WS1.Range(WS1.Cells(Selection(1).Row, 1), WS1.Cells(Selection(Selection.Count).Row, 45)).Copy
            WS1.Range(WS1.Cells(S1_A_LASTROW + 1, 1), WS1.Cells(S1_A_LASTROW + 1, 45)).PasteSpecial

            Dim S1_決定ETA_COL As Long    'シート1の1行目から"決定ETA"を探す
            S1_決定ETA_COL = WorksheetFunction.Match("決定ETA", .Rows("1:1"), 0)

            Dim S1_決定輸入港_COL As Long    'シート1の1行目から"決定輸入港"を探す
            S1_決定輸入港_COL = WorksheetFunction.Match("決定輸入港", .Rows("1:1"), 0)

            WS1.Range(WS1.Cells(S1_A_LASTROW + 1, S1_決定ETA_COL), WS1.Cells(S1_A_LASTROW + 1, S1_決定輸入港_COL)).ClearContents

            Dim S1_IVNO_COL As Long    'シート1の1行目から"INVOICE No"を探す
            S1_IVNO_COL = WorksheetFunction.Match("INVOICE No", .Rows("1:1"), 0)

            Dim S1_済_COL As Long    'シート1の1行目から"済"を探す
            S1_済_COL = WorksheetFunction.Match("済", .Rows("1:1"), 0)

            WS1.Range(WS1.Cells(S1_A_LASTROW + 1, S1_IVNO_COL), WS1.Cells(S1_A_LASTROW + 1, S1_済_COL)).ClearContents

            Dim S1_発注書No_COL As Long    'シート1の1行目から"発注書No"を探す
            S1_発注書No_COL = WorksheetFunction.Match("発注書No", .Rows("1:1"), 0)

            Dim S1_発行No_COL As Long    'シート1の1行目から"発行No"を探す
            S1_発行No_COL = WorksheetFunction.Match("発行No", .Rows("1:1"), 0)

            Dim S1_発注書No_CELL As String
            S1_発注書No_CELL = WS1.Cells(S1_A_LASTROW + 1, S1_発注書No_COL)

            Dim S1_発行No_CELL As String
            S1_発行No_CELL = WS1.Cells(S1_A_LASTROW + 1, S1_発行No_COL)

            .Range(.Cells(1, 1), .Cells(1, S1_LASTCOLUMN_1)).AutoFilter Field:=S1_発注書No_COL, _
                                                                        Criteria1:=S1_発注書No_CELL

            .Range(.Cells(1, 1), .Cells(1, S1_LASTCOLUMN_1)).AutoFilter Field:=S1_発行No_COL, _
                                                                        Criteria1:=S1_発行No_CELL

            Dim S1_数量_COL As Long    'シート1の1行目から"数量"を探す
            S1_数量_COL = WorksheetFunction.Match("数量", .Rows("1:1"), 0)

            WS1.Range(WS1.Cells(S1_A_LASTROW + 1, S1_数量_COL), WS1.Cells(S1_A_LASTROW + 1, S1_数量_COL)).Activate

            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
                   , AllowFiltering:=True    'プロテクト
            .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能
            Application.ScreenUpdating = True
            AppActivate Application.Caption
            Unload UserForm1
            '        MsgBox "処理終了"
            MsgBox "数量変更してください。決定ETA、決定輸入港も入力"
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