Option Explicit

'Sub 個人売上集計()
'
'Dim WS1 As Worksheet
'Set WS1 = ThisWorkbook.Worksheets("個人売上")
'
'Dim WS2 As Worksheet
'Set WS2 = ThisWorkbook.Worksheets("貼付")
'
'Dim WS3 As Worksheet
'Set WS3 = ThisWorkbook.Worksheets("営業リスト")
'
'Dim 売上明細 As String
'売上明細 = "C:\Users\toyocase\Desktop\日報_AO\DATA\売上明細.xlsx"
'
'Dim AO得意先マスタ As String
'AO得意先マスタ = "C:\Users\toyocase\Desktop\日報_AO\DATA\得意先マスタ.xlsx"
'
'Dim AO商品マスタ As String
'AO商品マスタ = "C:\Users\toyocase\Desktop\日報_AO\DATA\商品マスタ.xlsx"
'
'Dim 竜王得意先マスタ As String
'竜王得意先マスタ = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)\AOマスタ\得意先マスタ.xlsx"
'
'Dim 竜王商品マスタ As String
'竜王商品マスタ = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)\AOマスタ\商品マスタ.xlsx"
'
'Dim 本社得意先マスタ As String
'本社得意先マスタ = "\\192.168.1.9\業務\竜王在庫\AOマスタ(移動削除ファイル名変更禁止)\得意先マスタ.xlsx"
'
'Dim 本社商品マスタ As String
'本社商品マスタ = "\\192.168.1.9\業務\竜王在庫\AOマスタ(移動削除ファイル名変更禁止)\商品マスタ.xlsx"
'
'Dim AWS1 As Worksheet
'
'Dim FOR1 As Long
'
'Dim WS2_AC_LR As Long
'
'Dim 日報_AO As String
'日報_AO = "C:\Users\toyocase\Desktop\日報_AO\"
'
'If Dir(売上明細) = "" Then
'
'        MsgBox 売上明細 & vbCrLf & _
'               "が存在しません"
'
'               Exit Sub
'    Else
'
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
'WS2.Cells.ClearContents
'
'Workbooks.Open Filename:= _
'        売上明細, ReadOnly:=True
'
'End If
'
'Workbooks("売上明細.xlsx").Activate
'
'Set AWS1 = ActiveWorkbook.Worksheets("sheet1")
'
'AWS1.Cells.Copy
'WS2.Cells.PasteSpecial
'
'Application.CutCopyMode = False
'Workbooks("売上明細.xlsx").Close savechanges:=False
'
'WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row
'
'WS2.Cells(1, 47).Formula = "売上月"
'
'For FOR1 = 2 To WS2_AC_LR
'WS2.Cells(FOR1, 47).Value = Mid(WS2.Cells(FOR1, 1), 5, 2) * 1
'
'If WorksheetFunction.CountIf(WS3.Range("A:A"), WS2.Cells(FOR1, 19)) = 0 Then
'
'MsgBox "営業マスタにない営業です"
'
'Application.ScreenUpdating = True
'WS2.Cells(FOR1, 19).Activate
'Application.Calculation = xlCalculationAutomatic
'
'Exit Sub
'End If
'
'Next FOR1
'
'Application.Calculation = xlCalculationAutomatic
'
'  Application.DisplayAlerts = False
' ActiveWorkbook.SaveAs Filename:=日報_AO & "個人売上日報"
' Application.DisplayAlerts = True
'
' WS1.Copy
'
'Set AWS1 = ActiveWorkbook.Worksheets("個人売上")
'AWS1.Cells.Copy
'AWS1.Cells.PasteSpecial Paste:=xlPasteValues
'
'Application.DisplayAlerts = False
'
'With WorksheetFunction
'  ActiveWorkbook.SaveAs Filename:= _
' 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_個人売上日報"
'
'AWS1.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'        日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") _
'        & "_個人売上日報" & ".pdf", Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'        False
'End With
'
'ActiveWorkbook.Close savechanges:=False
'Application.DisplayAlerts = True
'
'WS1.Activate
'
'FileCopy AO得意先マスタ, 竜王得意先マスタ
'FileCopy AO商品マスタ, 竜王商品マスタ
'FileCopy AO得意先マスタ, 本社得意先マスタ
'FileCopy AO商品マスタ, 本社商品マスタ
'
'Application.ScreenUpdating = True
'
'MsgBox "更新完了"
'
'End Sub