Sub 仕入データ変換()

    Dim WS20 As Worksheet
    Set WS20 = ThisWorkbook.Worksheets("仕入伝票")

    WS20.Copy

    Dim AWS As Worksheet
    Set AWS = ActiveWorkbook.Worksheets("仕入伝票")

    AWS.Cells.Copy
    AWS.Cells.PasteSpecial Paste:=xlPasteValues

    Dim AWS_AC_LR As Long
    AWS_AC_LR = AWS.Cells(AWS.Rows.Count, 1).End(xlUp).Row

    Dim AWS_DC_LR As Long
    AWS_DC_LR = WorksheetFunction.CountA(Range("D1:D100")) - _
                WorksheetFunction.CountBlank(Range("D1:D100"))

    If AWS_AC_LR <> AWS_DC_LR Then

        AWS.Rows(AWS_DC_LR + 1 & ":" & AWS_AC_LR).Delete

    End If

    AWS_AC_LR = AWS.Cells(AWS.Rows.Count, 1).End(xlUp).Row

    Dim FOR1 As Long

    For FOR1 = 2 To AWS_AC_LR
        If Mid(AWS.Cells(FOR1, 24), 1) <> "'" Then
            AWS.Cells(FOR1, 24).Formula = "'" & AWS.Cells(FOR1, 24)
        End If
        'SendKeys "{F2}", True
        'SendKeys "{ENTER}", True

    Next FOR1

    AWS.OLEObjects.Delete

End Sub