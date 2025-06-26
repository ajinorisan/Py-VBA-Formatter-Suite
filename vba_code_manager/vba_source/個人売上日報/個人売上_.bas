Option Explicit

Sub 個人売上集計()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("個人売上")

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("貼付")

    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("営業リスト")

    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("59期-個人")

    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("ネット売上")

    If WS2.Visible = False Then
        WS2.Visible = True
    End If

    Dim 売上明細 As String
    売上明細 = ThisWorkbook.Path & "\DATA\担当者・得意先別売上.csv"

    Dim AO得意先マスタ As String
    AO得意先マスタ = ThisWorkbook.Path & "\DATA\得意先マスタ.xlsx"

    Dim AO商品マスタ As String
    AO商品マスタ = ThisWorkbook.Path & "\DATA\商品マスタ.xlsx"

    Dim 竜王得意先マスタ As String
    竜王得意先マスタ = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)\AOマスタ\得意先マスタ.xlsx"

    Dim 竜王商品マスタ As String
    竜王商品マスタ = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)\AOマスタ\商品マスタ.xlsx"

    Dim 本社得意先マスタ As String
    本社得意先マスタ = "\\192.168.1.9\業務\竜王在庫\AOマスタ(移動削除ファイル名変更禁止)\得意先マスタ.xlsx"

    Dim 本社商品マスタ As String
    本社商品マスタ = "\\192.168.1.9\業務\竜王在庫\AOマスタ(移動削除ファイル名変更禁止)\商品マスタ.xlsx"

    Dim 企画用商品マスタ As String
    企画用商品マスタ = "\\192.168.1.218\Product\08アラジンマスター\商品マスタ.xlsx"


    Dim AWS1 As Worksheet

    Dim FOR1 As Long

    Dim WS2_AC_LR As Long

    Dim 日報_AO As String
    日報_AO = ThisWorkbook.Path & "\"

    If Dir(売上明細) = "" Then

        MsgBox 売上明細 & vbCrLf & _
               "が存在しません"

        Exit Sub
    Else

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        WS2.Cells.ClearContents

        Workbooks.Open Filename:= _
                       売上明細, ReadOnly:=True

    End If

    '.Activate

    Set AWS1 = Workbooks("担当者・得意先別売上.csv").Worksheets("担当者・得意先別売上")

    Dim AWS1_AC_LR As Long
    AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row

    AWS1.Range("A:D").NumberFormatLocal = "@"

    With WorksheetFunction
        For FOR1 = 2 To AWS1_AC_LR - 1
            AWS1.Cells(FOR1, 1).Formula = .Text(AWS1.Cells(FOR1, 1), "000000")
            AWS1.Cells(FOR1, 3).Formula = .Text(AWS1.Cells(FOR1, 3), "00000000")
            AWS1.Cells(FOR1, 5).Value = Mid(AWS1.Cells(FOR1, 5), 5, 2)
        Next FOR1
    End With

    AWS1.Cells.Copy
    WS2.Cells.PasteSpecial

    Application.CutCopyMode = False
    Workbooks("担当者・得意先別売上.csv").Close SaveChanges:=False

    WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row

    WS2.Rows(WS2_AC_LR & ":" & WS2_AC_LR).Delete


    WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row
    'WS2.Cells(1, 47).Formula = "売上月"

    For FOR1 = 2 To WS2_AC_LR
        'WS2.Cells(FOR1, 5).Value = Mid(WS2.Cells(FOR1, 5), 5, 2) * 1
        If WS2.Cells(FOR1, 1) = "009997" Then
            WS2.Cells(FOR1, 1) = "009999"
        End If

        If WS2.Cells(FOR1, 1) = "009996" Then
            WS2.Cells(FOR1, 1) = "009999"
        End If

        'If WS2.Cells(FOR1, 1) = "000257" Then
        'WS2.Cells(FOR1, 1) = "000015"
        'End If

        If (WorksheetFunction.CountIf(WS3.Range("A:A"), WS2.Cells(FOR1, 1)) = 0) And (WS2.Cells(FOR1, 6) <> 0) Then

            MsgBox "営業マスタにない営業です"

            Application.ScreenUpdating = True
            WS2.Cells(FOR1, 2).Activate
            Application.Calculation = xlCalculationAutomatic

            Exit Sub
        End If

    Next FOR1

    WS2.Visible = False
    WS3.Visible = False
    WS4.Visible = False
    WS5.Visible = False

   'tiveWindow.FreezePanes = True

    'Application.Calculation = xlCalculationAutomatic

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=日報_AO & "個人売上日報"
    Application.DisplayAlerts = True

    WS1.Copy

    Set AWS1 = ActiveWorkbook.Worksheets("個人売上")
    AWS1.Cells.Copy
    AWS1.Cells.PasteSpecial Paste:=xlPasteValues

    Application.DisplayAlerts = False

    With WorksheetFunction
        ActiveWorkbook.SaveAs Filename:= _
                              日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_個人売上日報"

        AWS1.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                                 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") _
                                 & "_個人売上日報" & ".pdf", Quality:=xlQualityStandard, _
                                 IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
                                 False
    End With

    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True

    WS1.Activate

    FileCopy AO得意先マスタ, 竜王得意先マスタ
    FileCopy AO商品マスタ, 竜王商品マスタ
    FileCopy AO得意先マスタ, 本社得意先マスタ
    FileCopy AO商品マスタ, 本社商品マスタ
    On Error Resume Next
    FileCopy AO商品マスタ, 企画用商品マスタ

    Workbooks.Open 企画用商品マスタ

    Dim 企画用商品マスタシート As Worksheet
    Set 企画用商品マスタシート = ActiveWorkbook.Worksheets("sheet1")

    Application.ScreenUpdating = True
    企画用商品マスタシート.Cells(2, 3).Select

    'If ActiveWindow.FreezePanes = True Then
    'ActiveWindow.FreezePanes = False
    'End If
    '
    'ActiveWindow.FreezePanes = True

    ActiveWorkbook.Close SaveChanges:=True
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "更新完了"

End Sub
