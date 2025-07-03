Sub データ読込()

    Dim 元ファイル As Variant

    元ファイル = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")
    If 元ファイル = False Then Exit Sub
    Workbooks.Open 元ファイル

    Application.ScreenUpdating = False

    Dim AWB As Workbook
    Set AWB = ActiveWorkbook

    Dim AWS1 As Worksheet
    Set AWS1 = ActiveWorkbook.Worksheets(1)

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("データ貼付用")

    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("MENU")

    If WS1.AutoFilterMode = True Then
        WS1.AutoFilterMode = False
    End If

    AWS1.Cells.Copy
    WS1.Cells.PasteSpecial

    Application.DisplayAlerts = False

    AWB.Close
    Application.DisplayAlerts = True

    元ファイル = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")
    If 元ファイル = False Then GoTo オートフィット
    Workbooks.Open 元ファイル

    Dim AWB2 As Workbook
    Set AWB2 = ActiveWorkbook

    Dim AWS2 As Worksheet
    Set AWS2 = ActiveWorkbook.Worksheets(1)

    Set AWB2 = ActiveWorkbook
    Set AWS2 = ActiveWorkbook.Worksheets(1)
    'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
    Dim AWS2_AC_LR As Long
    AWS2_AC_LR = AWS2.Cells(AWS2.Rows.Count, 1).End(xlUp).Row

    AWS2.Range(AWS2.Cells(2, 1), AWS2.Cells(AWS2_AC_LR, 73)).Copy

    Dim WS1_AC_LR As Long
    WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

    WS1.Cells(WS1_AC_LR + 1, 1).PasteSpecial

    Application.DisplayAlerts = False

    AWB2.Close
    Application.DisplayAlerts = True

    オートフィット:

    With WS1.Cells
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    WS1.Cells.EntireColumn.AutoFit

    WS3.Activate

    Application.ScreenUpdating = True

End Sub