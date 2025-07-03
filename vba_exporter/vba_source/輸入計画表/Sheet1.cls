Private Sub CommandButton1_Click()

    'Application.WindowState = xlMinimized
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call 再計算

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    'Application.WindowState = xlMaximized

    MsgBox "更新完了"

End Sub