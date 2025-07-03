
Option Explicit

Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm1.Show
End Sub


Private Sub CommandButton2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ws2 As Worksheet

    Set ws2 = ThisWorkbook.Worksheets("予約")

    Call toggle_sheet_protection(ws2, "un_protect")

    Call toggle_screen_update(False)

    With ws2
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    ws2.AutoFilterMode = False



    Dim ws2_4_row_lastcol As Long
    ws2_4_row_lastcol = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

    ws2.Range(ws2.Cells(4, 1), ws2.Cells(4, ws2_4_row_lastcol)).AutoFilter

    Call toggle_sheet_protection(ws2, "protect")
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")

    Dim syohincd_col As Long
    syohincd_col = 2

    Dim nyuryoku_kaishi_row As Long
    nyuryoku_kaishi_row = 5

    Dim nyuryoku_kaishi_col As Long
    nyuryoku_kaishi_col = 13

    Dim ws2_lastrow_c_col As Long
    ws2_lastrow_c_col = ws2.Cells(ws1.Rows.Count, "C").End(xlUp).Row

    Application.EnableEvents = False
    Call toggle_screen_update(False)

    Dim i As Long

    If Target.Column = syohincd_col And Target.Row >= nyuryoku_kaishi_row And Target.Row <= ws2_lastrow_c_col - 1 And Selection.CountLarge = 1 Then

        Dim yoyakusuu As Variant
        yoyakusuu = ws1.Cells(Target.Row, 17).Value

        If yoyakusuu > 0 Then

            Call toggle_sheet_protection(ws2, "un_protect")

            For i = nyuryoku_kaishi_col To ws2_lastrow_c_col
                If ws2.Cells(Target.Row(), i) = "" Then

                    ws2.Columns(i).Hidden = True
                Else
                    ws2.Columns(i).Hidden = False
                End If


            Next i

            Call toggle_sheet_protection(ws2, "protect")

        End If

    End If

    Call toggle_screen_update(True)
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_Deactivate()
    Unload UserForm1
End Sub

