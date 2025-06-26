Sub 得意先別売上集計()

Dim TIME As Single
TIME = Timer   '時間計測開始

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim WS1 As Worksheet
Set WS1 = ThisWorkbook.Worksheets("得意先担当者別")

Dim WS2 As Worksheet
Set WS2 = ThisWorkbook.Worksheets("今期")

Dim WS3 As Worksheet
Set WS3 = ThisWorkbook.Worksheets("前期")

Dim WS4 As Worksheet
Set WS4 = ThisWorkbook.Worksheets("計算")

Dim WS2_B_LR As Long
WS2_B_LR = WS2.Cells(WS2.Rows.Count, 2).End(xlUp).Row

Dim WS3_B_LR As Long
WS3_B_LR = WS3.Cells(WS3.Rows.Count, 2).End(xlUp).Row

Dim 売上明細 As String
売上明細 = ThisWorkbook.Path & "\DATA\担当者・得意先別売上.csv"

Workbooks.Open Filename:= _
        売上明細, ReadOnly:=True

        Dim AWS1 As Worksheet

        Set AWS1 = Workbooks("担当者・得意先別売上.csv").Worksheets(1)

  AWS1.Cells.Copy
  WS2.Cells.PasteSpecial
  
  Application.DisplayAlerts = False
  Workbooks("担当者・得意先別売上.csv").Close savechanges:=False
  Application.DisplayAlerts = True

WS4.Cells.ClearContents

WS2.Range(WS2.Cells(2, 1), WS2.Cells(WS2_B_LR, 6)).Copy
WS4.Range("A1").PasteSpecial Paste:=xlPasteValues

Dim WS4_B_LR As Long
WS4_B_LR = WS4.Cells(WS4.Rows.Count, 2).End(xlUp).Row

WS3.Range(WS3.Cells(2, 1), WS3.Cells(WS3_B_LR, 6)).Copy
WS4.Cells(WS4_B_LR + 1, 1).PasteSpecial Paste:=xlPasteValues

Application.CutCopyMode = False

If WS4.AutoFilterMode = True Then
WS4.AutoFilterMode = False
End If

WS4_B_LR = WS4.Cells(WS4.Rows.Count, 2).End(xlUp).Row

WS4.Range("A1").AutoFilter Field:=6, Criteria1:=0

WS4.Rows("1:" & WS4_B_LR).EntireRow.Delete

WS4.AutoFilterMode = False

WS4.Range("A:F").RemoveDuplicates Columns:=Array(1, 3), Header:=xlNo

WS4_B_LR = WS4.Cells(WS4.Rows.Count, 2).End(xlUp).Row

WS4.Range(WS4.Cells(1, 1), WS4.Cells(WS4_B_LR, 6)) _
               .Sort Key1:=WS4.Cells(1, 3), order1:=xlAscending
                
WS4_B_LR = WS4.Cells(WS4.Rows.Count, 2).End(xlUp).Row

If WS1.AutoFilterMode = True Then
WS1.AutoFilterMode = False
End If

WS1.Range("B3:C6002").ClearContents
WS1.Range("T3:V6002").ClearContents
WS1.Range("A3:V6002") _
              .Sort Key1:=WS1.Range("A2"), order1:=xlAscending

Dim WS1_B_LR As Long
WS1_B_LR = WS1.Cells(WS1.Rows.Count, 2).End(xlUp).Row
              
Dim FOR1 As Long

For FOR1 = 3 To 6002 Step 3
WS1.Range(WS1.Cells(FOR1, 5), WS1.Cells(FOR1 + 1, 10)).ClearContents
WS1.Range(WS1.Cells(FOR1, 12), WS1.Cells(FOR1 + 1, 17)).ClearContents
Next FOR1

With WorksheetFunction

For FOR1 = 1 To WS4_B_LR

WS1.Cells(WS1_B_LR + 1, 2).Value = WS4.Cells(FOR1, 3).Value
WS1.Cells(WS1_B_LR + 1, 3).Value = WS4.Cells(FOR1, 4).Value
WS1.Cells(WS1_B_LR + 1, 20).Value = WS4.Cells(FOR1, 2).Value

Dim NO As Long

NO = 5
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

NO = 5
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 2 Then
NO = 6
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))
End If

NO = 6
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 3 Then
NO = 7
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))
End If

NO = 7
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 4 Then
NO = 8
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 8
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 5 Then
NO = 9
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 9
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 6 Then
NO = 10
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 10
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 7 Then
NO = 12
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 12
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 8 Then
NO = 13
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 13
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 9 Then
NO = 14
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 14
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 10 Then
NO = 15
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 15
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 11 Then
NO = 16
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 16
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 12 Then
NO = 17
WS1.Cells(WS1_B_LR + 1, NO).Value = _
.SumIfs(WS2.Range("F:F"), WS2.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS2.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 17
WS1.Cells(WS1_B_LR + 2, NO).Value = _
.SumIfs(WS3.Range("F:F"), WS3.Range("C:C"), WS1.Cells(WS1_B_LR + 1, 2), _
WS3.Range("B:B"), WS1.Cells(WS1_B_LR + 1, 20), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

WS1.Cells(WS1_B_LR + 2, 2).Value = WS4.Cells(FOR1, 3).Value
WS1.Cells(WS1_B_LR + 2, 3).Value = WS4.Cells(FOR1, 4).Value
WS1.Cells(WS1_B_LR + 2, 20).Value = WS4.Cells(FOR1, 2).Value

WS1.Cells(WS1_B_LR + 3, 2).Value = WS4.Cells(FOR1, 3).Value
WS1.Cells(WS1_B_LR + 3, 3).Value = WS4.Cells(FOR1, 4).Value
WS1.Cells(WS1_B_LR + 3, 20).Value = WS4.Cells(FOR1, 2).Value

WS1_B_LR = WS1.Cells(WS1.Rows.Count, 2).End(xlUp).Row

Next FOR1

Dim 得意先マスタ As String
得意先マスタ = ThisWorkbook.Path & "\DATA\得意先マスタ.xlsx"

Workbooks.Open Filename:= _
        得意先マスタ, ReadOnly:=True

Set AWS1 = Workbooks("得意先マスタ.xlsx").Worksheets("sheet1")

Dim AWS1_AC_LR As Long

AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row

For FOR1 = 2 To AWS1_AC_LR
AWS1.Cells(FOR1, 1).Value = .Clean(AWS1.Cells(FOR1, 1)) * 1
AWS1.Cells(FOR1, 40).Value = .Clean(AWS1.Cells(FOR1, 40))

Next FOR1

WS1.Calculate

WS1.Range("W:W").Delete

WS1_B_LR = WS1.Cells(WS1.Rows.Count, 2).End(xlUp).Row

For FOR1 = 3 To WS1_B_LR

WS1.Cells(FOR1, 21).Value = .Index(AWS1.Cells, .Match(WS1.Cells(FOR1, 2), AWS1.Range("A:A"), 0), 40)
On Error Resume Next
If (WS1.Cells(FOR1, 4) = "今期実績") And (WS1.Cells(FOR1, 19) <> 0) Then

WS1.Cells(FOR1, 23).Value = WS1.Cells(FOR1, 19)
End If
On Error GoTo 0

WS1.Range("W:W").Replace What:="#DIV/0!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False '#DIV/0!を空白に置き換え

Next FOR1

'ここからオカシイ
'WS1.Calculate
'Application.Wait Now() + TimeValue("00:00:05")



For FOR1 = 3 To WS1_B_LR Step 3
If WS1.Cells(FOR1, 23) <> "" Then

WS1.Cells(FOR1, 22).Value = .Rank(WS1.Cells(FOR1, 23), WS1.Range("W3:W6002"), 0)
WS1.Cells(FOR1 + 1, 22).Value = .Rank(WS1.Cells(FOR1, 23), WS1.Range("W3:W6002"), 0)
WS1.Cells(FOR1 + 2, 22).Value = .Rank(WS1.Cells(FOR1, 23), WS1.Range("W3:W6002"), 0)

End If
Next FOR1

WS1.Range("W:W").Delete

WS1.Range("A2:V2").AutoFilter

Workbooks("得意先マスタ.xlsx").Close savechanges:=False

WS1.Range("T6003:T6065").ClearContents
WS1.Range("C6003:C6065").ClearContents
WS4.Range("A:B").Copy
WS4.Range("H1").PasteSpecial

WS4.Range("H:I").RemoveDuplicates Columns:=1, Header:=xlNo

Dim WS4_H_LR As Long

WS4_H_LR = WS4.Cells(WS4.Rows.Count, 8).End(xlUp).Row

WS4.Range(WS4.Cells(1, 8), WS4.Cells(WS4_H_LR, 9)) _
               .Sort Key1:=WS4.Cells(1, 8), order1:=xlAscending
               
WS1.Cells(6003, 3).Value = WS4.Cells(1, 9)
WS1.Cells(6003 + 1, 3).Value = WS4.Cells(1, 9)
WS1.Cells(6003 + 2, 3).Value = WS4.Cells(1, 9)
WS1.Cells(6003, 20).Value = WS4.Cells(1, 9)
WS1.Cells(6003 + 1, 20).Value = WS4.Cells(1, 9)
WS1.Cells(6003 + 2, 20).Value = WS4.Cells(1, 9)

Dim WS1_C_LR As Long

WS1_C_LR = WS1.Cells(WS1.Rows.Count, 3).End(xlUp).Row
               
For FOR1 = 2 To WS4_H_LR

WS1.Cells(WS1_C_LR + 1, 3).Value = WS4.Cells(FOR1, 9)
WS1.Cells(WS1_C_LR + 2, 3).Value = WS4.Cells(FOR1, 9)
WS1.Cells(WS1_C_LR + 3, 3).Value = WS4.Cells(FOR1, 9)
WS1.Cells(WS1_C_LR + 1, 20).Value = WS4.Cells(FOR1, 9)
WS1.Cells(WS1_C_LR + 2, 20).Value = WS4.Cells(FOR1, 9)
WS1.Cells(WS1_C_LR + 3, 20).Value = WS4.Cells(FOR1, 9)

WS1_C_LR = WS1.Cells(WS1.Rows.Count, 3).End(xlUp).Row

Next FOR1

For FOR1 = 6003 To 6063 Step 3
WS1.Range(WS1.Cells(FOR1, 5), WS1.Cells(FOR1 + 1, 10)).ClearContents
WS1.Range(WS1.Cells(FOR1, 12), WS1.Cells(FOR1 + 1, 17)).ClearContents
Next FOR1

For FOR1 = 6003 To 6063 Step 3

NO = 5
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

NO = 5
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 2 Then
NO = 6
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))
End If

NO = 6
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 3 Then
NO = 7
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))
End If

NO = 7
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 4 Then
NO = 8
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 8
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 5 Then
NO = 9
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 9
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 6 Then
NO = 10
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 10
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 7 Then
NO = 12
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 12
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 8 Then
NO = 13
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 13
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 9 Then
NO = 14
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 14
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 10 Then
NO = 15
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 15
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 11 Then
NO = 16
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 16
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

If WS1.Cells(1, 1) >= 12 Then
NO = 17
WS1.Cells(FOR1, NO).Value = _
.SumIfs(WS2.Range("F:F"), _
WS2.Range("B:B"), WS1.Cells(FOR1, 3), _
WS2.Range("E:E"), Mid(WS2.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))
End If

NO = 17
WS1.Cells(FOR1 + 1, NO).Value = _
.SumIfs(WS3.Range("F:F"), _
WS3.Range("B:B"), WS1.Cells(FOR1 + 1, 3), _
WS3.Range("E:E"), Mid(WS3.Cells(2, 5), 1, 4) + 1 & "0" & WS1.Cells(2, NO))

Next FOR1

End With

WS1.Copy
'Dim AWS1 As Worksheet
Set AWS1 = ActiveWorkbook.Worksheets("得意先担当者別")
AWS1.Cells(1, 4).Copy
AWS1.Cells(1, 4).PasteSpecial Paste:=xlPasteValues

Dim 日報_AO As String
日報_AO = ThisWorkbook.Path & "\"

Application.DisplayAlerts = False
With WorksheetFunction
  ActiveWorkbook.SaveAs Filename:= _
 日報_AO & Year(Date) & .Text(Month(Date), "00") & .Text(Day(Date), "00") & "_得意先別担当者別対比表.xlsx"
End With
Application.DisplayAlerts = True

ActiveWorkbook.Close savechanges:=False

WS1.Activate

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "更新完了" & vbCrLf & "処理時間は " & Round(Timer - TIME, 2) & " 秒です。"

End Sub