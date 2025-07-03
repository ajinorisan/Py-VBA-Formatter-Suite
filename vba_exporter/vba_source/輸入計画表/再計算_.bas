Sub 再計算()

    '商品マスタの再計算

    Dim WS3 As Worksheet

    Set WS3 = ThisWorkbook.Sheets("商品マスタ")

    Dim WS3_A列_LASTROW As Long

    WS3_A列_LASTROW = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

    '※後で取る
    ' Application.Calculation = xlCalculationManual

    Dim FOR1 As Long

    For FOR1 = 2 To WS3_A列_LASTROW
        WS3.Cells(FOR1, 1).Value = StrConv(WS3.Cells(FOR1, 1), vbNarrow)
        WS3.Cells(FOR1, 1).Value = WorksheetFunction.Clean(WS3.Cells(FOR1, 1))
        'WS3.Cells(FOR1, 8).Value = (WS3.Cells(FOR1, 10) * WS3.Cells(FOR1, 11) * WS3.Cells(FOR1, 12)) / 1000000

    Next FOR1

    '※後で取る
    'Application.Calculation = xlCalculationAutomatic

    '入力シートの形式を整える

    Dim WS1 As Worksheet

    Set WS1 = ThisWorkbook.Sheets("入力")

    Dim WS1_A列_LASTROW As Long

    If WS1.AutoFilterMode = True Then
        WS1.AutoFilterMode = False
        WS1.Range("A2:S2").AutoFilter field:=19, Criteria1:=""
    End If

    WS1.Range("A2").EntireRow.RowHeight = 50

    WS1_A列_LASTROW = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

    WS1.Range("D3:D" & WS1_A列_LASTROW).ClearContents
    WS1.Range("F3:J" & WS1_A列_LASTROW).ClearContents
    WS1.Range("O3:O" & WS1_A列_LASTROW).ClearContents

    WS1.Range("D1").Copy

    WS1.Range("D3:D" & WS1_A列_LASTROW).PasteSpecial

    WS1.Range("F1:J1").Copy

    WS1.Range("F3:J" & WS1_A列_LASTROW).PasteSpecial

    WS1.Range("O1").Copy

    WS1.Range("O3:O" & WS1_A列_LASTROW).PasteSpecial

    WS1.Range("F:G").NumberFormatLocal = "$#,##0.00;[赤]-$#,##0.00"

    For FOR1 = 3 To WS1_A列_LASTROW

        If InStr(WS1.Cells(FOR1, 1).Value, "レーベル") > 0 Then
            WS1.Cells(FOR1, 6).NumberFormatLocal = "\#,##0;\-#,##0"
            WS1.Cells(FOR1, 7).NumberFormatLocal = "\#,##0;\-#,##0"
        End If

    Next FOR1

    WS1.Calculate

    WS1.Columns("A:S").EntireColumn.AutoFit

    WS1.Range("A:S").Borders.LineStyle = xlLineStyleNone

    For FOR1 = 3 To WS1_A列_LASTROW

        WS1.Cells(FOR1, 2).Value = StrConv(WS1.Cells(FOR1, 2), vbNarrow)

        If WS1.Cells(FOR1, 1) <> WS1.Cells(FOR1 + 1, 1) Or _
            WS1.Cells(FOR1, 14) <> WS1.Cells(FOR1 + 1, 14) Then

            WS1.Range(WS1.Cells(FOR1, 1), WS1.Cells(FOR1, 19)). _
            Borders(xlEdgeBottom).LineStyle = xlContinuous

        End If

    Next FOR1

    Application.CutCopyMode = False

    '計画表データ作成

    Dim WS4 As Worksheet

    Set WS4 = ThisWorkbook.Sheets("計画表データ")

    WS4.Visible = True

    Dim WS4_A列_LASTROW As Long

    WS4_A列_LASTROW = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row

    If WS4_A列_LASTROW >= 2 Then

        WS4.Rows("2:" & WS4_A列_LASTROW).Delete

    End If

    WS1.Range("N3:N" & WS1_A列_LASTROW).SpecialCells(xlCellTypeVisible).Copy

    WS4.Range("A2").PasteSpecial xlPasteValues

    For FOR1 = 3 To WS1_A列_LASTROW

        If WS1.Cells(FOR1, 2) <> "" Then
            WS1.Cells(FOR1, 20).Value = WS1.Cells(FOR1, 2)
        Else
            WS1.Cells(FOR1, 20).Value = WS1.Cells(FOR1, 3)
        End If

    Next FOR1

    WS1.Range("T3:T" & WS1_A列_LASTROW).SpecialCells(xlCellTypeVisible).Copy

    WS4.Range("B2").PasteSpecial xlPasteValues

    WS1.Columns(20).Delete

    WS1.Range("E3:E" & WS1_A列_LASTROW).SpecialCells(xlCellTypeVisible).Copy

    WS4.Range("C2").PasteSpecial xlPasteValues

    For FOR1 = 2 To WS4_A列_LASTROW

        If WorksheetFunction.CountIf(WS4.Range(WS4.Cells(2, 2), WS4.Cells(FOR1, 2)), WS4.Cells(FOR1, 2)) >= 2 Then
            WS4.Cells(FOR1, 4).Value = WS4.Cells(FOR1, 1)
            WS4.Cells(FOR1, 5).Value = WS4.Cells(FOR1, 2)
            WS4.Cells(FOR1, 6).Value = WS4.Cells(FOR1, 3)
        End If

        If WorksheetFunction.CountIf(WS4.Range(WS4.Cells(2, 2), WS4.Cells(FOR1, 2)), WS4.Cells(FOR1, 2)) >= 3 Then
            WS4.Cells(FOR1, 7).Value = WS4.Cells(FOR1, 1)
            WS4.Cells(FOR1, 8).Value = WS4.Cells(FOR1, 2)
            WS4.Cells(FOR1, 9).Value = WS4.Cells(FOR1, 3)
        End If

        If WorksheetFunction.CountIf(WS4.Range(WS4.Cells(2, 2), WS4.Cells(FOR1, 2)), WS4.Cells(FOR1, 2)) >= 4 Then
            WS4.Cells(FOR1, 10).Value = WS4.Cells(FOR1, 1)
            WS4.Cells(FOR1, 11).Value = WS4.Cells(FOR1, 2)
            WS4.Cells(FOR1, 12).Value = WS4.Cells(FOR1, 3)
        End If

    Next FOR1

    WS4.Columns("A:L").EntireColumn.AutoFit
    WS4.Range("A:A").NumberFormatLocal = "yyyy/m/d;@"
    WS4.Range("D:D").NumberFormatLocal = "yyyy/m/d;@"
    WS4.Range("G:G").NumberFormatLocal = "yyyy/m/d;@"
    WS4.Range("J:J").NumberFormatLocal = "yyyy/m/d;@"
    WS4.Range("C:C").NumberFormatLocal = "#,##0_ ;[赤]-#,##0"
    WS4.Range("F:F").NumberFormatLocal = "#,##0_ ;[赤]-#,##0"
    WS4.Range("I:I").NumberFormatLocal = "#,##0_ ;[赤]-#,##0"
    WS4.Range("L:L").NumberFormatLocal = "#,##0_ ;[赤]-#,##0"

    '在庫表更新用処理

    Sheets("計画表データ").Copy
    ActiveWorkbook.Worksheets.Add.Name = "更新履歴"

    Dim AWS2 As Worksheet
    Set AWS2 = ActiveWorkbook.Worksheets("更新履歴")

    AWS2.Range("C:C,E:E,G:G,I:I").NumberFormatLocal = "m""月""d""日"";@"
    AWS2.Range("D:D,F:F,H:H,J:J").NumberFormatLocal = "#,##0_ ;[赤]-#,##0 "
    AWS2.Columns("K:K").EntireColumn.Hidden = True

    Application.PrintCommunication = False
    With AWS2.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    AWS2.PageSetup.PrintArea = ""
    '    Application.PrintCommunication = False
    With AWS2.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        '        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    '    Application.PrintCommunication = True

    ChDir "\\192.168.2.2\share\竜王共有ファイル\輸入品入荷表\輸入計画データ"
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:= _
    "\\192.168.2.2\share\竜王共有ファイル\輸入品入荷表\輸入計画データ\輸入計画データ.xlsx", FileFormat:= _
    xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    Application.DisplayAlerts = True

    WS4.Visible = False

End Sub