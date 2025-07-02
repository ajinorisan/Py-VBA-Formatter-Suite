Private Sub CommandButton1_Click()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Call データ処理
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Call データ読込
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton3_Click()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Call AO用データ作成
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton4_Click()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    Call 納品先コード振分
    Application.ScreenUpdating = True
End Sub