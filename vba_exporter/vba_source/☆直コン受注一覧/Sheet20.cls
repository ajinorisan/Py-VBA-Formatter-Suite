Private Sub CommandButton1_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call 仕入データ変換

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub