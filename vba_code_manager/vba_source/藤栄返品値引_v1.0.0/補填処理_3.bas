Option Explicit

Sub 補填処理_3()

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


    WS6.Columns("M:Z").Delete
    Dim WS6_AC_LR As Long
    WS6_AC_LR = WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Row

    Dim i As Long

    With WorksheetFunction
        For i = 2 To WS6_AC_LR
            On Error Resume Next
            WS6.Cells(i, 13).Value = .Clean(.VLookup(WS6.Cells(i, 1) * 1, WS4.Columns("A:F"), 3, False)) * 1
            WS6.Cells(i, 13).Value = Format(WS6.Cells(i, 13), "'00000000")
            WS6.Cells(i, 14).Value = .VLookup(WS6.Cells(i, 1) * 1, WS4.Columns("A:F"), 4, False)
            WS6.Cells(i, 15).Value = .Clean(.VLookup(WS6.Cells(i, 1) * 1, WS4.Columns("A:F"), 5, False)) * 1
            WS6.Cells(i, 15).Value = Format(WS6.Cells(i, 15), "'00000000")
            WS6.Cells(i, 16).Value = .VLookup(WS6.Cells(i, 1) * 1, WS4.Columns("A:F"), 6, False)
            On Error GoTo 0

            On Error GoTo SYOHINERROR
            WS6.Cells(i, 17).Value = .VLookup(WS6.Cells(i, 7) * 1, WS3.Columns("A:B"), 2, False)
            WS6.Cells(i, 18).Value = .VLookup(WS6.Cells(i, 7) * 1, WS3.Columns("A:D"), 4, False)
            WS6.Cells(i, 19).Value = WS6.Cells(i, 18) - WS6.Cells(i, 11) * 1
            On Error GoTo 0
            WS6.Cells(i, 20).Value = Format(WS6.Cells(i, 3) & "/20", "YYMMDD")
        Next i
    End With

    For i = 2 To WS6_AC_LR
        If WS6.Cells(i, 14) = "" Then
            WS6.Activate
            MsgBox "A列見直しして問題なければ、" & i & "行の拠点マスタを登録してください"
            WS4.Activate
            Exit Sub
        ElseIf WS6.Cells(i, 17) = "" Then
            WS6.Activate
            MsgBox i & "行の商品マスタの商品コードを登録してください"
            WS3.Activate
            Exit Sub
        ElseIf WS6.Cells(i, 18) = "" Then
            WS6.Activate
            MsgBox i & "行の商品マスタの単価を登録してください"
            WS3.Activate
            Exit Sub
        End If


    Next i

    WS7.Rows("2:3000").Delete
    Dim WS7_EC_LR As Long
    WS7_EC_LR = WS7.Cells(WS7.Rows.Count, 5).End(xlUp).Row

    For i = 2 To WS6_AC_LR

        WS7.Cells(WS7_EC_LR + 1, 5).Value = WS6.Cells(i, 13)
        WS7.Cells(WS7_EC_LR + 1, 6).Value = WS6.Cells(i, 14)
        WS7.Cells(WS7_EC_LR + 1, 7).Value = WS6.Cells(i, 15)
        WS7.Cells(WS7_EC_LR + 1, 8).Value = WS6.Cells(i, 16)

        WS7.Cells(WS7_EC_LR + 1, 9).Formula = "'00000998"
        'WS7.Cells(WS7_EC_LR + 1, 10).Value = WS6.Cells(i, 6)
        WS7.Cells(WS7_EC_LR + 1, 11).Formula = "'03"
        WS7.Cells(WS7_EC_LR + 1, 12).Value = WS6.Cells(i, 17)
        WS7.Cells(WS7_EC_LR + 1, 18).Value = -WS6.Cells(i, 10)
        WS7.Cells(WS7_EC_LR + 1, 19).Value = WS6.Cells(i, 18)
        WS7.Cells(WS7_EC_LR + 1, 20).Value = WS7.Cells(WS7_EC_LR + 1, 18) * WS7.Cells(WS7_EC_LR + 1, 19)
        WS7.Cells(WS7_EC_LR + 1, 21).Value = WS6.Cells(i, 20)
        WS7.Cells(WS7_EC_LR + 1, 25).Value = WS6.Cells(i, 5) & " " & WS6.Cells(i, 6)

        WS7.Cells(WS7_EC_LR + 2, 5).Value = WS6.Cells(i, 13)
        WS7.Cells(WS7_EC_LR + 2, 6).Value = WS6.Cells(i, 14)
        WS7.Cells(WS7_EC_LR + 2, 7).Value = WS6.Cells(i, 15)
        WS7.Cells(WS7_EC_LR + 2, 8).Value = WS6.Cells(i, 16)

        WS7.Cells(WS7_EC_LR + 2, 9).Formula = "'00000998"
        'WS7.Cells(WS7_EC_LR + 1, 10).Value = WS6.Cells(i, 6)
        WS7.Cells(WS7_EC_LR + 2, 11).Formula = "'04"
        WS7.Cells(WS7_EC_LR + 2, 12).Value = WS6.Cells(i, 17)
        WS7.Cells(WS7_EC_LR + 2, 18).Value = WS6.Cells(i, 10)
        WS7.Cells(WS7_EC_LR + 2, 19).Value = WS6.Cells(i, 19)
        WS7.Cells(WS7_EC_LR + 2, 20).Value = WS7.Cells(WS7_EC_LR + 2, 18) * WS7.Cells(WS7_EC_LR + 2, 19)
        WS7.Cells(WS7_EC_LR + 2, 21).Value = WS6.Cells(i, 20)
        WS7.Cells(WS7_EC_LR + 2, 25).Value = WS6.Cells(i, 5) & " " & WS6.Cells(i, 6)

        WS7_EC_LR = WS7_EC_LR + 2
    Next i

    With WorksheetFunction
        If -.Sum(WS7.Range("T:T")) <> .Sum(WS6.Range("L:L")) Then
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
    AWB_NAME = Mid(WS7.Cells(2, 1), 3, 4) & "藤栄補填"

    ' 保存先とファイル名を指定（例：Desktopに保存）
    SAVE_PATH = Environ("USERPROFILE") & "\Desktop\" & AWB_NAME

    ' ワークブックを保存
    ActiveWorkbook.SaveAs Filename:=SAVE_PATH

    ' ワークブックを閉じる
    ActiveWorkbook.Close

    WS2.Copy

    Dim AWS As Worksheet
    Set AWS = ActiveWorkbook.Worksheets("出力")
    AWS.Cells.Copy
    AWS.Cells.PasteSpecial xlPasteValues

    AWB_NAME = Mid(WS7.Cells(2, 1), 3, 4) & "藤栄返品補填明細"

    ' 保存先とファイル名を指定（例：Desktopに保存）
    SAVE_PATH = Environ("USERPROFILE") & "\Desktop\" & AWB_NAME

    ' ワークブックを保存
    ActiveWorkbook.SaveAs Filename:=SAVE_PATH

    Application.DisplayAlerts = True

    MsgBox "補填処理終了"
    Exit Sub
SYOHINERROR:
    Dim WS3_LR_AC As Long
    WS3_LR_AC = WS3.Cells(Rows.Count, 1).End(xlUp).Row

    Dim ITEM_CODE As Double
    ITEM_CODE = WS6.Cells(i, 7) * 1

    Dim ITEM_RESULT As Range
    Set ITEM_RESULT = WS3.Columns("A").Find(ITEM_CODE, LookIn:=xlValues, LookAt:=xlWhole)

    If ITEM_RESULT Is Nothing Then
        ' 見つからない場合、新しい行を一番下に追加

        WS3.Cells(WS3_LR_AC + 1, 1).Value = WS6.Cells(i, 7) * 1


    End If

    Resume Next


End Sub

