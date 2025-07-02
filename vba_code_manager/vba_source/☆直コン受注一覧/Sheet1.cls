Private Sub CommandButton21_Click()
    UserForm1.Show
    AppActivate Application.Caption
End Sub

Private Sub Worksheet_Activate()
    UserForm1.Show
    AppActivate Application.Caption
End Sub

Private Sub Worksheet_Deactivate()
    Unload UserForm1
    AppActivate Application.Caption
End Sub