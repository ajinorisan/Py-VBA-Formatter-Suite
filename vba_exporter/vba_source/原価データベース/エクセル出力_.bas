Option Explicit
Sub エクセル出力()

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("原価計算書")

    WS2.Copy

    Dim AWS As Worksheet
    Set AWS = ActiveWorkbook.Worksheets("原価計算書")

    AWS.Unprotect
    AWS.Range("O4").Copy
    AWS.Range("O4").PasteSpecial Paste:=xlPasteValues
    AWS.Range("R1:S1").Copy
    AWS.Range("R1:S1").PasteSpecial Paste:=xlPasteValues
    AWS.Range("S4:S5").Copy
    AWS.Range("S4:S5").PasteSpecial Paste:=xlPasteValues
    AWS.Range("A7:S11").Copy
    AWS.Range("A7:S11").PasteSpecial Paste:=xlPasteValues
    AWS.Range("S13").Copy
    AWS.Range("S13").PasteSpecial Paste:=xlPasteValues
    AWS.Range("A13:R48").Copy
    AWS.Range("A13:R48").PasteSpecial Paste:=xlPasteValues
    AWS.Range("A50:G51").Copy
    AWS.Range("A50:G51").PasteSpecial Paste:=xlPasteValues
    AWS.Range("A53:G57").Copy
    AWS.Range("A53:G57").PasteSpecial Paste:=xlPasteValues
    AWS.Range("S53:S57").Copy
    AWS.Range("S53:S57").PasteSpecial Paste:=xlPasteValues
    AWS.Range("L53:P57").Copy
    AWS.Range("L53:P57").PasteSpecial Paste:=xlPasteValues
    AWS.Range("A60:S64").Copy
    AWS.Range("A60:S64").PasteSpecial Paste:=xlPasteValues
    AWS.Range("R51:S51").Copy
    AWS.Range("R51:S51").PasteSpecial Paste:=xlPasteValues
    'AWS.Range("M51:Q51").Copy
    'AWS.Range("M51:Q51").PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    Dim OBJ As OLEObject
    For Each OBJ In AWS.OLEObjects
        If OBJ.progID = "Forms.CommandButton.1" Then OBJ.Delete
    Next OBJ

    Application.DisplayAlerts = False

    Dim AWBN As String
    AWBN = Application.ActiveWorkbook.Name

    Dim 保存フルパス As String
    保存フルパス = "C:\Users\TOYOC-304\Desktop\TEMP\" & AWBN & ".xlsx"

    Dim AWB As Workbook
    Set AWB = ActiveWorkbook

    AWB.SaveAs Filename:=保存フルパス

    Application.DisplayAlerts = True
    ActiveWindow.WindowState = xlMinimized

End Sub