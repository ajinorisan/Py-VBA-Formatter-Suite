Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Cancel = True

Dim WS3 As Worksheet
Set WS3 = ThisWorkbook.Worksheets("アラジン取込用(受注)")

Dim WS4 As Worksheet
Set WS4 = ThisWorkbook.Worksheets("アラジン取込用(売上)")

If (WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row) _
And WS3.Cells(2, 9) = WS4.Cells(2, 10) Then

WS3.Range("2:3000").Delete

Dim WS3_AC_LR As Long
WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row


WS3.Cells(WS3_AC_LR + 1, 1).Formula = WS4.Cells(Target.Row(), 1)
WS3.Cells(WS3_AC_LR + 1, 2).Formula = WS4.Cells(Target.Row(), 1)
WS3.Cells(WS3_AC_LR + 1, 4).Formula = WS4.Cells(Target.Row(), 5)
WS3.Cells(WS3_AC_LR + 1, 6).Formula = "80013313"
WS3.Cells(WS3_AC_LR + 1, 8).Formula = WS4.Cells(Target.Row(), 9)
WS3.Cells(WS3_AC_LR + 1, 9).Formula = WS4.Cells(Target.Row(), 10)
WS3.Cells(WS3_AC_LR + 1, 10).Formula = "'01"
WS3.Cells(WS3_AC_LR + 1, 11).Formula = WS4.Cells(Target.Row(), 12)
WS3.Cells(WS3_AC_LR + 1, 13).Formula = WS4.Cells(Target.Row(), 14)
WS3.Cells(WS3_AC_LR + 1, 17).Formula = WS4.Cells(Target.Row(), 18)
WS3.Cells(WS3_AC_LR + 1, 18).Formula = WS4.Cells(Target.Row(), 19)

If WS4.Cells(Target.Row(), 26) <> "" Then

WS3.Cells(WS3_AC_LR + 1, 23).Formula = WS4.Cells(Target.Row(), 25) & "(注)" & WS4.Cells(Target.Row(), 26)
Else
WS3.Cells(WS3_AC_LR + 1, 23).Formula = WS4.Cells(Target.Row(), 25)
End If

WS4.Rows(Target.Row()).Delete

Else

'Dim WS3_AC_LR As Long
WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row


WS3.Cells(WS3_AC_LR + 1, 1).Formula = WS4.Cells(Target.Row(), 1)
WS3.Cells(WS3_AC_LR + 1, 2).Formula = WS4.Cells(Target.Row(), 1)
WS3.Cells(WS3_AC_LR + 1, 4).Formula = WS4.Cells(Target.Row(), 5)
WS3.Cells(WS3_AC_LR + 1, 6).Formula = "80013313"
WS3.Cells(WS3_AC_LR + 1, 8).Formula = WS4.Cells(Target.Row(), 9)
WS3.Cells(WS3_AC_LR + 1, 9).Formula = WS4.Cells(Target.Row(), 10)
WS3.Cells(WS3_AC_LR + 1, 10).Formula = "'01"
WS3.Cells(WS3_AC_LR + 1, 11).Formula = WS4.Cells(Target.Row(), 12)
WS3.Cells(WS3_AC_LR + 1, 13).Formula = WS4.Cells(Target.Row(), 14)
WS3.Cells(WS3_AC_LR + 1, 17).Formula = WS4.Cells(Target.Row(), 18)
WS3.Cells(WS3_AC_LR + 1, 18).Formula = WS4.Cells(Target.Row(), 19)
If WS4.Cells(Target.Row(), 26) <> "" Then

WS3.Cells(WS3_AC_LR + 1, 23).Formula = WS4.Cells(Target.Row(), 25) & "(注)" & WS4.Cells(Target.Row(), 26)
Else
WS3.Cells(WS3_AC_LR + 1, 23).Formula = WS4.Cells(Target.Row(), 25)
End If
WS4.Rows(Target.Row()).Delete

'Dim WSH As Object
'    Set WSH = CreateObject("WScript.Shell")
'    WSH.Popup "転記完了", 1, "メッセージ", vbInformation
'    Set WSH = Nothing

End If
End Sub

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Application.ScreenUpdating = True
'End Sub