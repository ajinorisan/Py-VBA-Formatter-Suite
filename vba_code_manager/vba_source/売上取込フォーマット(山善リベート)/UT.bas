Option Explicit



Public Function 最終行取得(ワークシート As Worksheet, 列 As Long)
    最終行取得 = ワークシート.Cells(ワークシート.Rows.Count, 列).End(xlUp).Row
End Function

Public Sub 行削除(ワークシート As Worksheet, 開始行 As Long, 最終行 As Long)
    ワークシート.Rows(開始行 & ":" & 最終行).Delete
End Sub

Public Sub 自動更新停止(マクロ一覧消去用 As Long)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
End Sub

Public Sub 自動更新再開(マクロ一覧消去用 As Long)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Cursor = xlDefault
End Sub

Public Sub 上書確認停止(マクロ一覧消去用 As Long)
    Application.DisplayAlerts = False
End Sub

Public Sub 上書確認再開(マクロ一覧消去用 As Long)
    Application.DisplayAlerts = True
End Sub
