Option Explicit

Private Sub Worksheet_Deactivate()

    Dim ws4 As Worksheet
    Set ws4 = ThisWorkbook.Worksheets("入出庫明細")
    Call clean_vlue(ws4, 7)

End Sub