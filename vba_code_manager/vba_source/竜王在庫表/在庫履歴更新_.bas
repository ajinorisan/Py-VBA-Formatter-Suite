Option Explicit

Public Sub 在庫履歴更新()

    Dim rireki_path As String
    rireki_path = "\\192.168.1.218\業務\在庫履歴\"


    Dim aw_path As String
    aw_path = ActiveWorkbook.Path

    Call toggle_screen_update(False)
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ChDir rireki_path
    ActiveWorkbook.SaveAs Filename:=rireki_path & Format(Date, "yymmdd") & "竜王在庫表.xlsm"
    
    ActiveWorkbook.SaveAs Filename:= _
                          aw_path & "\竜王在庫表.xlsm", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False



  
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call toggle_screen_update(True)
    MsgBox "在庫履歴更新終了"

End Sub

'Format(Date, "yymmdd") & "簡易在庫表.xlsx"
'Sub 在庫履歴更新()
'
''保存先のフルパス
'Dim 竜王在庫履歴保存先パス As String
''竜王在庫履歴保存先パス = "\\192.168.2.2\share\竜王共有ファイル\在庫表\在庫履歴\20170301~竜王在庫表\"
'竜王在庫履歴保存先パス = "\\192.168.1.218\業務\在庫履歴\"
'Dim AWP As String
'AWP = ActiveWorkbook.Path
'
''画面変更を止める
'Application.ScreenUpdating = False
''他のマクロ停止
'Application.EnableEvents = False
'    '保存先指定
'    ChDir 竜王在庫履歴保存先パス
'    '上書きするかの警告を非表示
'    Application.DisplayAlerts = False
'    On Error Resume Next
'    'コピーを書込みパスワード付で保存
'    With WorksheetFunction
'    ActiveWorkbook.SaveAs fileName:= _
      '        竜王在庫履歴保存先パス & Year(Date) & .text(Month(Date), "00") & .text(Day(Date), "00") & "竜王在庫表.xlsm", FileFormat:= _
      '        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
'        End With
'        On Error GoTo 0
'
'        ActiveWorkbook.SaveAs fileName:= _
          '        AWP & "\竜王在庫表_AO.xlsm", FileFormat:= _
          '        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
'
'        '他のマクロ開始
'        Application.EnableEvents = True
'        '警告を再表示
'        Application.DisplayAlerts = True
'        '画面変更を開始
'                Application.ScreenUpdating = True
'                MsgBox "在庫履歴更新終了"
'
'End Sub
