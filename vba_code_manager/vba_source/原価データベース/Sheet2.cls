Option Explicit

Private Sub CommandButton21_Click()
Application.ScreenUpdating = False
Call エクセル出力
Application.ScreenUpdating = True

ThisWorkbook.Activate
Application.Windows("原価データベース.xlsm").Activate

End Sub