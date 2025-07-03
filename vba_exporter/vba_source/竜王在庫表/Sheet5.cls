Option Explicit

Private Sub Worksheet_Deactivate()

    Dim ws5 As Worksheet
    Set ws5 = ThisWorkbook.Worksheets("棚卸原票")
    Call clean_vlue(ws5, 3)

End Sub
