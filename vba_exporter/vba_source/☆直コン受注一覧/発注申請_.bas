Option Explicit

Sub 発注申請()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")

    Dim WS14 As Worksheet
    Set WS14 = ThisWorkbook.Worksheets("貼付")

    Dim WS21 As Worksheet
    Set WS21 = ThisWorkbook.Worksheets("発注申請用")

    With WS1

        .Unprotect

        .Cells.EntireColumn.Hidden = False

        Dim S1_A_LASTROW As Long    'シート1のA列の最後のデータ行取得
        S1_A_LASTROW = .Cells(Rows.Count, 1).End(xlUp).Row

        Dim S1_LASTCOLUMN_1 As Long    'シート1の1行目の最後のデータ列取得
        S1_LASTCOLUMN_1 = .Cells(1, Columns.Count).End(xlToLeft).Column

        .Cells.Copy
        WS14.Cells.PasteSpecial xlPasteValues

        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True    'プロテクト

        .EnableSelection = xlNoRestrictions    'ロックしたセルも参照可能

        '   MsgBox "貼付完了"

    End With

    Dim WS14_A_LASTROW As Long    'シート14のA列の最後のデータ行取得
    WS14_A_LASTROW = WS14.Cells(WS14.Rows.Count, 1).End(xlUp).Row

    WS21.Rows("2:3000").Delete

    WS21.Range("A1:N1").Copy
    WS21.Range(WS21.Cells(2, 1), WS21.Cells(WS14_A_LASTROW, 14)).PasteSpecial

    WS21.Cells(WS14_A_LASTROW + 1, 14) = _
    WorksheetFunction.Sum(WS21.Range(WS21.Cells(2, 14), WS21.Cells(WS14_A_LASTROW, 14)))

    WS21.Cells(WS14_A_LASTROW + 1, 7) = _
    WorksheetFunction.Sum(WS21.Range(WS21.Cells(2, 7), WS21.Cells(WS14_A_LASTROW, 7)))

    WS21.Cells(WS14_A_LASTROW + 1, 9) = _
    WorksheetFunction.Sum(WS21.Range(WS21.Cells(2, 9), WS21.Cells(WS14_A_LASTROW, 9)))

    WS21.Columns("A:N").EntireColumn.AutoFit

    WS21.Copy

    ActiveSheet.Cells.Copy
    ActiveSheet.Cells.PasteSpecial Paste:=xlPasteValues
    ActiveSheet.Columns("A:N").EntireColumn.AutoFit

End Sub
