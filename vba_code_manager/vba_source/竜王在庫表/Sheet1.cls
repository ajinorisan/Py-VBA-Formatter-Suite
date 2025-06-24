Option Explicit

Private Sub CommandButton4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm1.Show
End Sub

Private Sub CommandButton5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserForm2.Show
End Sub

Private Sub Worksheet_Deactivate()

    Unload UserForm1
    Unload UserForm2

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = ThisWorkbook.Worksheets("在庫表")
    Set ws2 = ThisWorkbook.Worksheets("予約")

    Dim yoyaku_col As Long
    yoyaku_col = 17

    Dim yoyaku_start_col As Long
    yoyaku_start_col = 13

    Dim ws1_lastrow_d_col As Long
    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

    Dim i As Long

    Application.EnableEvents = False
    Call toggle_screen_update(False)

    If Target.Column = yoyaku_col And Target.Row >= 5 And Target.Row <= ws1_lastrow_d_col And Selection.CountLarge = 1 Then
        If Target.Value <> 0 Then
            ws2.Activate
            Call toggle_sheet_protection(ws2, "un_protect")


            Dim ws2_4_row_lastcol As Long
            ws2_4_row_lastcol = ws2.Cells(4, ws2.Columns.Count).End(xlToLeft).Column

            For i = yoyaku_start_col To ws2_4_row_lastcol
                If ws2.Cells(Target.Row(), i) = "" Then
                    
                    ws2.Columns(i).Hidden = True
                Else
                    ws2.Columns(i).Hidden = False
                End If
            Next i



            ws2.Cells(Target.Row, 1).Select

            Application.Goto ws2.Cells(Target.Row, 1), True

            Call toggle_sheet_protection(ws2, "protect")

        End If
    End If

    Call toggle_screen_update(True)
    Application.EnableEvents = True
End Sub


Private Sub CommandButton_MouseDown(btn_type As Integer)
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("在庫表")

    Call toggle_sheet_protection(ws1, "un_protect")

    Dim ws1_4_row_lastcol As Long
    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

    Call toggle_screen_update(False)

    Dim i As Long

    For i = 1 To ws1_4_row_lastcol
        If ws1.Cells(1, i) = btn_type Or ws1.Cells(1, i) = 0 Then
            ws1.Columns(i).Hidden = False
        Else
            ws1.Columns(i).Hidden = True
        End If
    Next i

    ws1.Rows("1:1").Hidden = True

    Call toggle_sheet_protection(ws1, "protect")

    Call toggle_screen_update(True)
End Sub

Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton_MouseDown (1)
End Sub

Private Sub CommandButton2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CommandButton_MouseDown (2)
End Sub


Private Sub CommandButton3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("在庫表")

    Call toggle_sheet_protection(ws1, "un_protect")

    Dim ws1_lastrow_d_col As Long
    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row

    Dim ws1_4_row_lastcol As Long
    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column

    Call toggle_screen_update(False)

    With ws1
        .Rows.Hidden = False
        .Columns.Hidden = False
    End With

    ws1.AutoFilterMode = False

    ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_4_row_lastcol)).AutoFilter

    ws1.Rows("1:1").Hidden = True

    Call toggle_sheet_protection(ws1, "protect")

    Call toggle_screen_update(True)
End Sub

'Private Sub CommandButton1_Click()    '出荷用
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    Dim i As Long
'
'    For i = 1 To ws1_4_row_lastcol
'        If ws1.Cells(1, i) = 1 Or ws1.Cells(1, i) = 0 Then
'            ws1.Columns(i).Hidden = False
'        Else
'            ws1.Columns(i).Hidden = True
'        End If
'    Next i
'
'    ws1.Rows("1:1").Hidden = True
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'    Call ThisWorkbook.toggle_screen_update(True)
'
'End Sub
'
'
'Private Sub CommandButton2_Click()    '入出庫
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    Dim i As Long
'
'    For i = 1 To ws1_4_row_lastcol
'        If ws1.Cells(1, i) = 2 Or ws1.Cells(1, i) = 0 Then
'            ws1.Columns(i).Hidden = False
'        Else
'            ws1.Columns(i).Hidden = True
'        End If
'    Next i
'
'    ws1.Rows("1:1").Hidden = True
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'    Call ThisWorkbook.toggle_screen_update(True)
'
'End Sub

'Private Sub CommandButton3_Click()    '全体表示
'
'    Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("在庫表")
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "un_protect")
'
'    Dim ws1_lastrow_d_col As Long
'    ws1_lastrow_d_col = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
'
'    Dim ws1_4_row_lastcol As Long
'    ws1_4_row_lastcol = ws1.Cells(4, ws1.Columns.Count).End(xlToLeft).Column
'
'    Call ThisWorkbook.toggle_screen_update(False)
'
'    With ws1
'        .Rows.Hidden = False
'        .Columns.Hidden = False
'    End With
'
'    ws1.AutoFilterMode = False
'
'    ws1.Range(ws1.Cells(4, 1), ws1.Cells(4, ws1_4_row_lastcol)).AutoFilter
'
'    ws1.Rows("1:1").Hidden = True
'
'    Call ThisWorkbook.toggle_sheet_protection(ws1, "protect")
'
'    Call ThisWorkbook.toggle_screen_update(True)
'
'End Sub
