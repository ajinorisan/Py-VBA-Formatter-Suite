Option Explicit

Private Sub Worksheet_Deactivate()

    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Worksheets("受注残")
    Call clean_vlue(ws3, 7)

End Sub
