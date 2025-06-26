Option Explicit

Sub 連続印刷()

    Dim WS8 As Worksheet
    Set WS8 = ThisWorkbook.Worksheets("作業指示書")

    Dim AP1 As String
    AP1 = Application.ActivePrinter

    Dim PN1 As String
    PN1 = "DocuCentre-IV C5575 on Ne"

    Dim PN2 As String
    Dim TRUE1 As Boolean
    Dim FOR1 As Long

    On Error Resume Next
    For FOR1 = 1 To 99
        PN2 = PN1 & Format(FOR1, "00:")

        Application.ActivePrinter = PN2
        If Application.ActivePrinter = PN2 Then
            TRUE1 = True
            Exit For
        End If

    Next FOR1
    On Error GoTo 0

    Dim IB1 As Long
    Dim IB2 As Long

    IB1 = Application.InputBox(Prompt:="最初の売上No入力", Type:=1)
    If IB1 = False Then GoTo 終了
    IB2 = Application.InputBox(Prompt:="最後の売上No入力", Type:=1)
    If IB2 = False Then GoTo 終了

    For FOR1 = IB1 To IB2

        WS8.Cells(3, 4).Value = FOR1
        WS8.Calculate
        WS8.PrintOut ActivePrinter:=PN2

    Next FOR1

    MsgBox "印刷完了 " & IB2 - IB1 + 1 & "枚"

終了:

    Application.ActivePrinter = AP1

End Sub