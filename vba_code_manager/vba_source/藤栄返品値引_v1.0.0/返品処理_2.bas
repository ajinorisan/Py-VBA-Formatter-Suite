Option Explicit

Sub 返品処理_2()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("貼付")
    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("出力")
    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("商品マスタ")
    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("拠点マスタ")
    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("返品データ加工")
    Dim WS6 As Worksheet
    Set WS6 = ThisWorkbook.Worksheets("補填データ加工")
    Dim WS7 As Worksheet
    Set WS7 = ThisWorkbook.Worksheets("売上取込フォーマット")

    WS5.Columns("S:Z").Delete
    Dim WS5_AC_LR As Long
    WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row

    Dim i As Long

    With WorksheetFunction
        For i = 2 To WS5_AC_LR
            On Error Resume Next
            WS5.Cells(i, 19).Value = .Clean(.VLookup(WS5.Cells(i, 1), WS4.Columns("A:F"), 3, False)) * 1
            WS5.Cells(i, 19).Value = Format(WS5.Cells(i, 19), "'00000000")
            WS5.Cells(i, 20).Value = .VLookup(WS5.Cells(i, 1), WS4.Columns("A:F"), 4, False)
            WS5.Cells(i, 21).Value = .Clean(.VLookup(WS5.Cells(i, 1), WS4.Columns("A:F"), 5, False)) * 1
            WS5.Cells(i, 21).Value = Format(WS5.Cells(i, 21), "'00000000")
            WS5.Cells(i, 22).Value = .VLookup(WS5.Cells(i, 1), WS4.Columns("A:F"), 6, False)
            On Error GoTo 0

            On Error GoTo SYOHINERROR
            WS5.Cells(i, 23).Value = .VLookup(WS5.Cells(i, 11) * 1, WS3.Columns("A:B"), 2, False)
            On Error GoTo 0
            WS5.Cells(i, 24).Value = Format(WS5.Cells(i, 7), "YYMMDD")
        Next i
    End With

    For i = 2 To WS5_AC_LR
        If WS5.Cells(i, 19) = "" Then
            WS5.Activate
            MsgBox "A列見直しして問題なければ、" & i & "行の拠点マスタを登録してください"
            WS4.Activate
            Exit Sub
        ElseIf WS5.Cells(i, 23) = "" Then
            WS5.Activate
            MsgBox i & "行の商品マスタを登録してください"
            WS3.Activate
            Exit Sub
        End If


    Next i

    WS7.Rows("2:3000").Delete
    Dim WS7_EC_LR As Long
    WS7_EC_LR = WS7.Cells(WS7.Rows.Count, 5).End(xlUp).Row

    For i = 2 To WS5_AC_LR

        WS7.Cells(WS7_EC_LR + 1, 5).Value = WS5.Cells(i, 19)
        WS7.Cells(WS7_EC_LR + 1, 6).Value = WS5.Cells(i, 20)
        WS7.Cells(WS7_EC_LR + 1, 7).Value = WS5.Cells(i, 21)
        WS7.Cells(WS7_EC_LR + 1, 8).Value = WS5.Cells(i, 22)

        WS7.Cells(WS7_EC_LR + 1, 9).Formula = "'00000999"
        WS7.Cells(WS7_EC_LR + 1, 10).Value = WS5.Cells(i, 8)
        WS7.Cells(WS7_EC_LR + 1, 11).Formula = "'02"
        WS7.Cells(WS7_EC_LR + 1, 12).Value = WS5.Cells(i, 23)
        WS7.Cells(WS7_EC_LR + 1, 18).Value = -WS5.Cells(i, 14)
        WS7.Cells(WS7_EC_LR + 1, 19).Value = WS5.Cells(i, 15)
        WS7.Cells(WS7_EC_LR + 1, 20).Value = WS5.Cells(i, 16)
        WS7.Cells(WS7_EC_LR + 1, 21).Value = WS5.Cells(i, 24)
        WS7.Cells(WS7_EC_LR + 1, 25).Value = WS5.Cells(i, 17)
        WS7_EC_LR = WS7_EC_LR + 1
    Next i

    With WorksheetFunction
        If .Sum(WS7.Range("T:T")) <> .Sum(WS5.Range("P:P")) Then
            MsgBox "合計金額が合いません"
            Exit Sub
        End If
    End With


    Dim IB As String

    IB = InputBox("日付入力" & vbLf & vbLf & "20240121")

    If IB = "" Then

        Exit Sub
    Else
        WS7.Cells(2, 1).Value = IB
        WS7_EC_LR = WS7.Cells(Rows.Count, 5).End(xlUp).Row
        WS7.Cells(2, 1).Copy WS7.Range(WS7.Cells(2, 1), WS7.Cells(WS7_EC_LR, 2))
        WS7.Cells(2, 1).Copy WS7.Range(WS7.Cells(2, 24), WS7.Cells(WS7_EC_LR, 24))
    End If
    WS7.Copy

    Application.DisplayAlerts = False
    Dim SAVE_PATH As String
    Dim AWB_NAME As String
    AWB_NAME = Mid(WS7.Cells(2, 1), 3, 4) & "藤栄返品"

    ' 保存先とファイル名を指定（例：Desktopに保存）
    SAVE_PATH = Environ("USERPROFILE") & "\Desktop\" & AWB_NAME

    ' ワークブックを保存
    ActiveWorkbook.SaveAs Filename:=SAVE_PATH

    ' ワークブックを閉じる
    ActiveWorkbook.Close

    Application.DisplayAlerts = True

    MsgBox "返品処理終了"

    Exit Sub

SYOHINERROR:
    Dim WS3_LR_AC As Long
    WS3_LR_AC = WS3.Cells(Rows.Count, 1).End(xlUp).Row

    Dim ITEM_CODE As Double
    ITEM_CODE = WS5.Cells(i, 11) * 1

    Dim ITEM_RESULT As Range
    Set ITEM_RESULT = WS3.Columns("A").Find(ITEM_CODE, LookIn:=xlValues, LookAt:=xlWhole)

    If ITEM_RESULT Is Nothing Then
        ' 見つからない場合、新しい行を一番下に追加

        WS3.Cells(WS3_LR_AC + 1, 1).Value = WS5.Cells(i, 11) * 1

    End If

    Resume Next


End Sub