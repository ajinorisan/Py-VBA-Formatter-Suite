Option Explicit

Sub 選択貼付()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("入力")
    Dim WS14 As Worksheet
    Set WS14 = ThisWorkbook.Worksheets("貼付")

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

        MsgBox "貼付完了"

    End With

End Sub