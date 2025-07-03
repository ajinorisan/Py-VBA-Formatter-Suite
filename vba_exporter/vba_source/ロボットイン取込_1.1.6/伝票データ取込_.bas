Option Explicit

Sub 伝票データ取込()
    Dim 元ファイル As Variant

    元ファイル = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")
    If 元ファイル = False Then Exit Sub
    Workbooks.Open 元ファイル

    Application.ScreenUpdating = False

    Dim AWB As Workbook
    Set AWB = ActiveWorkbook

    Dim AWS1 As Worksheet
    Set AWS1 = ActiveWorkbook.Worksheets(1)

    Dim WS7 As Worksheet
    Set WS7 = ThisWorkbook.Worksheets("貼付")

    WS7.Cells.ClearContents
    AWS1.Cells.Copy
    WS7.Cells.PasteSpecial

    Application.DisplayAlerts = False

    AWB.Close
    Application.DisplayAlerts = True

    Dim WS8 As Worksheet
    Set WS8 = ThisWorkbook.Worksheets("作業指示書")

    WS8.Activate

    Application.ScreenUpdating = True

End Sub