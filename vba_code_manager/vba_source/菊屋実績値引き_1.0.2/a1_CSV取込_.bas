Option Explicit

Public Sub a1_CSV取込()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode = True Then
            ws.AutoFilterMode = False
        End If
    Next ws

    Dim ws7 As Worksheet
    Set ws7 = ThisWorkbook.Worksheets("支払明細書")
    Dim ws11 As Worksheet
    Set ws11 = ThisWorkbook.Worksheets("実績値引明細書")

    On Error Resume Next
    ws7.Cells.ClearFormats
    ws11.Cells.ClearFormats
    On Error GoTo 0

    Dim last_row As Long
    last_row = get_last_row(ws7, "A")
    ws7.Rows("1:" & last_row).Delete

    Call import_csv_file("支払明細書")
    ws7.Cells.Style = "Normal"

    last_row = get_last_row(ws11, "A")
    ws11.Rows("1:" & last_row).Delete

    Call import_csv_file("実績値引明細書")
    ws11.Cells.Style = "Normal"

End Sub