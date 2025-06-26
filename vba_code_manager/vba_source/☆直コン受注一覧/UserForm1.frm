
Private Sub CommandButton1_Click() '通常時並び替え

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 通常時並替
  
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
  Unload UserForm1
End Sub

Private Sub CommandButton10_Click()
Call 分納
End Sub

Private Sub CommandButton11_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'Application.Calculation = xlCalculationAutomatic
Call 通常時並替

 Application.Calculation = xlCalculationAutomatic
 Application.Calculation = xlCalculationManual
  Call 直コン表作成
  
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
  Unload UserForm1
End Sub

Private Sub CommandButton12_Click()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 山善送付用明細作成

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
'  AppActivate Application.Caption
  Unload UserForm1
  
End Sub

Private Sub CommandButton13_Click()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic

Call 発注申請

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
'  AppActivate Application.Caption
  Unload UserForm1
End Sub

Private Sub CommandButton2_Click() 'コンテナ明細並び替え

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 売上コンテナ並替

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
  Unload UserForm1
End Sub

Private Sub CommandButton3_Click() '受注入力時並び替え

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 受注入力

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
Unload UserForm1
End Sub

Private Sub CommandButton4_Click() '発注時並び替え

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 発注入力

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
Unload UserForm1

End Sub
Private Sub CommandButton5_Click() 'ETA変更用並び替え

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call ETA変更

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
Unload UserForm1
End Sub

Private Sub CommandButton6_Click() 'INVOICE他入力

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call IV入力

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
 Unload UserForm1
End Sub

Private Sub CommandButton7_Click() '仕入入力用

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Call 仕入入力

Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True
  AppActivate Application.Caption
  Unload UserForm1
End Sub

Private Sub CommandButton8_Click() '行クリア
Call 行クリア
End Sub

Private Sub CommandButton9_Click()
Application.ScreenUpdating = False
Call 選択貼付
Application.ScreenUpdating = True
  AppActivate Application.Caption
  Unload UserForm1
End Sub