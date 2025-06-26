Option Explicit

Private Sub CommandButton1_Click()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False
Call 連続印刷
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = False
Call 伝票データ取込
Application.ScreenUpdating = True
End Sub