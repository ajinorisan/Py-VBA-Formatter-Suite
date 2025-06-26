Sub 例外処理()

'ここよりTWB例外処理

Dim WS1 As Worksheet
Set WS1 = ThisWorkbook.Worksheets("データ貼付用")


Dim WS1_AC_LR As Long
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

Dim FOR1 As Long
For FOR1 = 2 To WS1_AC_LR
If WorksheetFunction.CountIf(WS1.Cells(FOR1, 74), "TWB*") = 1 Then
Dim WS1_81C_BK As Long
Dim WS1_81C_BL As Long
Dim WS1_81C_RE As Long
WS1_81C_BK = WorksheetFunction.CountIf(WS1.Cells(FOR1, 81), "*ブラック*")

WS1_81C_BL = WorksheetFunction.CountIf(WS1.Cells(FOR1, 81), "*ブルー*")

WS1_81C_RE = WorksheetFunction.CountIf(WS1.Cells(FOR1, 81), "*レッド*")

If WS1_81C_BK = 1 Then

If Len(WS1.Cells(FOR1, 74)) = 6 Then
WS1.Cells(FOR1, 74).Value = "TWB-BK"
End If

If Len(WS1.Cells(FOR1, 74)) = 9 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 1)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-BK"
End If

If Len(WS1.Cells(FOR1, 74)) = 10 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 2)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-BK"
End If

End If

If WS1_81C_BL = 1 Then

If Len(WS1.Cells(FOR1, 74)) = 6 Then
WS1.Cells(FOR1, 74).Value = "TWB-BL"
End If

If Len(WS1.Cells(FOR1, 74)) = 9 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 1)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-BL"
End If

If Len(WS1.Cells(FOR1, 74)) = 10 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 2)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-BL"
End If

End If


If WS1_81C_RE = 1 Then

If Len(WS1.Cells(FOR1, 74)) = 6 Then
WS1.Cells(FOR1, 74).Value = "TWB-RE"
End If

If Len(WS1.Cells(FOR1, 74)) = 9 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 1)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-RE"
End If

If Len(WS1.Cells(FOR1, 74)) = 10 Then
WS1.Cells(FOR1, 77).Formula = Mid(WS1.Cells(FOR1, 74), 5, 2)
WS1.Cells(FOR1, 76).Value = WS1.Cells(FOR1, 76) / WS1.Cells(FOR1, 77)
WS1.Cells(FOR1, 74).Value = "TWB-RE"
End If

End If

End If

'BGH-DB例外処理
If WS1.Cells(FOR1, 74) = "BGH-DB-6P-IV" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-IV"
WS1.Cells(FOR1, 76).Value = 415
WS1.Cells(FOR1, 77).Value = 6
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "BGH-DB-3P-IV" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-IV"
WS1.Cells(FOR1, 76).Value = 530
WS1.Cells(FOR1, 77).Value = 3
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "BGH-DB-6P-PK" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-PK"
WS1.Cells(FOR1, 76).Value = 415
WS1.Cells(FOR1, 77).Value = 6
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "BGH-DB-3P-PK" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-PK"
WS1.Cells(FOR1, 76).Value = 530
WS1.Cells(FOR1, 77).Value = 3
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "BGH-DB-6P-BL" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-BL"
WS1.Cells(FOR1, 76).Value = 415
WS1.Cells(FOR1, 77).Value = 6
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "BGH-DB-3P-BL" Then
WS1.Cells(FOR1, 74).Formula = "BGH-DB-BL"
WS1.Cells(FOR1, 76).Value = 530
WS1.Cells(FOR1, 77).Value = 3
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

'COOL-ANB例外処理
If WS1.Cells(FOR1, 74) = "COOL-ANB-01-ANB2" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "COOL-ANB-01"
WS1.Cells(FOR1, 76).Value = 2480
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "COOL-ANB2-01"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 1760
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "COOL-ANB-02-ANB2" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "COOL-ANB-02"
WS1.Cells(FOR1, 76).Value = 2480
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "COOL-ANB2-02"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 1760
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "COOL-ANB-03-ANB2" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "COOL-ANB-03"
WS1.Cells(FOR1, 76).Value = 2480
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "COOL-ANB2-03"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 1760
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

'WRIS52-WFIS11例外処理
If WS1.Cells(FOR1, 74) = "WRIS52-WFIS11-ANBL" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS52-ANLBL"
WS1.Cells(FOR1, 76).Value = 2490
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2490
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "WRIS52-WFIS11-ANRE" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS52-ANRE"
WS1.Cells(FOR1, 76).Value = 2490
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2490
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "WRIS52-WFIS11-ANNV" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS52-ANNV"
WS1.Cells(FOR1, 76).Value = 2490
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2490
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

'WRIS51-WFIS11例外処理
If WS1.Cells(FOR1, 74) = "WRIS51-WFIS11-APRE" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS51-APRE"
WS1.Cells(FOR1, 76).Value = 2990
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2990
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "WRIS51-WFIS11-APGR" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS51-APGR"
WS1.Cells(FOR1, 76).Value = 2990
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2990
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If

If WS1.Cells(FOR1, 74) = "WRIS51-WFIS11-APNV" Then
WS1.Rows(FOR1 & ":" & FOR1).Copy
WS1.Rows(WS1_AC_LR + 1 & ":" & WS1_AC_LR + 1).PasteSpecial
WS1.Cells(FOR1, 74).Formula = "WRIS51-APNV"
WS1.Cells(FOR1, 76).Value = 2990
WS1.Cells(WS1_AC_LR + 1, 74).Formula = "WFIS11-NV"
WS1.Cells(WS1_AC_LR + 1, 76).Value = 2990
WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row
End If


Next FOR1

'ESTN-3P例外処理

For FOR1 = 2 To WS1_AC_LR

If (Mid(WS1.Cells(FOR1, 74), 1, 11) = "ESTN-CBM-3P") Or _
(Mid(WS1.Cells(FOR1, 74), 1, 11) = "ESTN-CBL-3P") Or _
(Mid(WS1.Cells(FOR1, 74), 1, 12) = "ESTN-CBLL-3P") Or _
(Mid(WS1.Cells(FOR1, 74), 1, 12) = "ESTN-CBUB-3P") Then

WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "-3P", "")
WS1.Cells(FOR1, 76) = WS1.Cells(FOR1, 76) / 3
WS1.Cells(FOR1, 77) = 3

End If

Next FOR1
'ESTN-3P例外処理終わり

'ﾒﾃﾞｨｱｸﾗﾌﾄ例外処理

For FOR1 = 2 To WS1_AC_LR

If WS1.Cells(FOR1, 74).Value Like "[WEB-B5*,WEB-BB*,WEB-CO*,WEB-DV*]" Then
If WS1.Cells(FOR1, 74).Value Like "*10P*" Then
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "10P-BK", "BK-10P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "10P-N", "N-10P")
Else
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "5P-BK", "BK-5P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "5P-N", "N-5P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "4P-BK", "BK-4P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "4P-N", "N-4P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "3P-BK", "BK-3P")
WS1.Cells(FOR1, 74) = Replace(WS1.Cells(FOR1, 74), "3P-N", "N-3P")
End If
End If

Next FOR1
'ﾒﾃﾞｨｱｸﾗﾌﾄ例外処理終わり

WS1.Range("A1:EZ1").AutoFilter

WS1.AutoFilter.Sort.SortFields.Clear
    WS1.AutoFilter.Sort.SortFields.Add2 Key:= _
        WS1.Range("E1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With WS1.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

WS1.AutoFilterMode = False
'COOL-ANB例外処理ここまで

'TWB例外処理ここまで

End Sub