
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
On Error Resume Next
Application.EnableEvents = False
Application.DisplayAlerts = False

Dim Filepath As String



Filepath = "\\192.168.1.218\業務\貿易\山善直コン\バックアップ\" _
& Format(Now, "yymmdd") & "_☆直コン受注一覧.xlsm"
ThisWorkbook.SaveAs Filename:=Filepath



Filepath = "\\192.168.1.218\Personal\yamazen\" _
& Format(Now, "yymmdd") & "_☆直コン受注一覧.xlsx"
ThisWorkbook.SaveAs Filename:=Filepath, FileFormat:=xlOpenXMLWorkbook

Filepath = "\\192.168.1.218\業務\貿易\山善直コン\☆直コン受注一覧.xlsm"
ThisWorkbook.SaveAs Filename:=Filepath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
'Dim Filepath As String


Application.DisplayAlerts = True
Application.EnableEvents = True
On Error GoTo 0

End Sub

Private Sub Workbook_Deactivate()
Unload UserForm1
'AppActivate Application.Caption
End Sub

