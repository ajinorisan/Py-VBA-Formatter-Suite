Option Explicit


'
'Private Sub CommandButton1_Click()
'Set S1 = ThisWorkbook.Worksheets("売上集計表")
'Set S2 = ThisWorkbook.Worksheets("今期DATA")
'Set S3 = ThisWorkbook.Worksheets("シリーズデータ")
'Set S4 = ThisWorkbook.Worksheets("商品分類MST")
'Set S6 = ThisWorkbook.Worksheets("作業")
'Set S7 = ThisWorkbook.Worksheets("PVT")
'Set S8 = ThisWorkbook.Worksheets("部門別商品別")
'
'
'部門別集計 = "C:\Users\toyocase\Desktop\日報用データ\売上集計_商品別_部門別用.csv"
' 日報場所 = "C:\Users\toyocase\Desktop\日報"
' 商品一覧CSV = "C:\Users\toyocase\Desktop\日報用データ\商品一覧.csv"
'
'Dim AWB As Workbook
'Dim AWBS As Worksheet
' Dim AWBN As String
'
' Dim AWS8 As Worksheet
' Dim AWS10 As Worksheet
' Dim AWS11 As Worksheet
'  Dim AWS12 As Worksheet
'   Dim AWS13 As Worksheet
'    Dim AWS14 As Worksheet
'
'Dim AWBS_AC_LR As Long
'Dim AWBS_LC_1R As Long
'Dim S1_BC_LR As Long
'Dim S2_BC_LR As Long
'Dim S6_AC_LR As Long
'Dim S4_AC_LR As Long
'Dim S4_BC_LR As Long
'Dim FOR1 As Long
'
' Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
'
''売上集計_商品別_部門別用.csv取込
'Workbooks.Open Filename:= _
'        部門別集計, ReadOnly:=True
'
'        Set AWB = ActiveWorkbook
'        AWB.ActiveSheet.Cells.Copy
'
' With S1
'
'        .Activate
'
'       .Cells.PasteSpecial Paste:=xlPasteValues
'       Application.CutCopyMode = False
'AWB.Close savechanges:=False
'
''手直し中データか判定
''If .Cells(1, 1) <> "月" Then
'
''1列目削除
'If .Cells(1, 2) = "売上集計表" Then
'.Rows("1:1").Delete Shift:=xlUp
'End If
'
''B列最終行取得
'S1_BC_LR = .Cells(Rows.Count, 2).End(xlUp).Row
'
'If .Cells(S1_BC_LR, 2).Value = "総合計" Then
'
'.Rows("1:1").AutoFilter
'
'    .Range("$A$1:$M$" & S1_BC_LR).AutoFilter Field:=2, Criteria1:=Array( _
'        "事業部<<合計>>", "総合計", "日付<<合計>>", "SYSTEM-9001"), Operator:=xlFilterValues
'
'    .Rows("2:" & S1_BC_LR).Delete Shift:=xlUp
'
'    .AutoFilterMode = False
'
'    End If
'
'    .AutoFilterMode = False
'
'    'いらない列削除後B列最終行再取得
'    S1_BC_LR = .Cells(Rows.Count, 2).End(xlUp).Row
'
'
'    'A列に売上月
' For FOR1 = 2 To S1_BC_LR
'
' If .Cells(FOR1, 1) = "" Then
'
' If .Cells(FOR1, 4) = "" Then
' .Cells(FOR1, 1).Value = Format(.Cells(FOR1, 2), "m")
' Else
' .Cells(FOR1, 1).Value = .Cells(FOR1 - 1, 1).Value
' End If
'
' End If
' Next FOR1
'
'     .Cells(1, 1).Formula = "月"
'
'  '売上日付データ行削除
' .Rows("1:1").AutoFilter
'
'     .Range("$A$1:$M$" & S1_BC_LR).AutoFilter Field:=3, Criteria1:="="
'     .Rows("2:" & S1_BC_LR).Delete Shift:=xlUp
'    .AutoFilterMode = False
'
''    End If
'
'  'いらない列削除後B列最終行再取得
'    S1_BC_LR = .Cells(Rows.Count, 2).End(xlUp).Row
'
'      For FOR1 = 2 To S1_BC_LR
'      .Cells(FOR1, 14).Value = Trim(.Cells(FOR1, 2))
'      Next FOR1
'
'      .Range("N2:N" & S1_BC_LR).Cut Destination:=.Range("B2:B" & S1_BC_LR)
'
'     S6.Columns("A:B").ClearContents
'
'     .Columns("B:B").Copy
'
'    End With
'
'    With S6
'
'    '仮商品MST作成
'
' .Columns("A:A").PasteSpecial Paste:=xlPasteValues
' .Range("$A$1:$A" & S1_BC_LR).RemoveDuplicates Columns:=1, Header:=xlNo
'
'    S6_AC_LR = .Cells(Rows.Count, 1).End(xlUp).Row
'
'    .Cells(2, 2).Formula = "=COUNTIF(商品分類MST!A:A,$A2)"
'    .Cells(2, 2).Copy .Range(.Cells(3, 2), .Cells(S6_AC_LR, 2))
'
'   .Calculate
''    Application.Calculation = xlCalculationAutomatic
''    Application.Calculation = xlCalculationManual
'
'   '商品分類MSTを増やす
'     For FOR1 = 2 To S6_AC_LR
'
'If .Cells(FOR1, 2) = 0 Then
'
' S4_AC_LR = S4.Cells(Rows.Count, 1).End(xlUp).Row
'
'' .Cells(FOR1, 3).Value = Trim(.Cells(FOR1, 1))
'' .Cells(FOR1, 3).Copy .Cells(FOR1, 1)
'
' .Cells(FOR1, 1).Copy S4.Cells(S4_AC_LR + 1, 1)
'
' End If
'
' Next FOR1
'
'    S4_AC_LR = S4.Cells(Rows.Count, 1).End(xlUp).Row
'
'End With
'
'    '新商品MST取り込み
'    With WorksheetFunction
'
'   If .CountA(S4.Range("A1:A" & S4_AC_LR)) <> _
'  .CountA(S4.Range("B1:B" & S4_AC_LR)) Then
'
'    Workbooks.Open Filename:= _
'        商品一覧CSV, ReadOnly:=True
'
'        Set AWB = ActiveWorkbook
'Set AWBS = ActiveWorkbook.ActiveSheet
'
' AWBS_AC_LR = AWBS.Cells(Rows.Count, 1).End(xlUp).Row
' AWBS_LC_1R = AWBS.Cells(1, Columns.Count).End(xlToLeft).Column
'
'S4_BC_LR = S4.Cells(Rows.Count, 2).End(xlUp).Row
'
'For FOR1 = S4_BC_LR + 1 To S4_AC_LR
''商品名
'S4.Cells(FOR1, 2).Value = _
'.Index(AWBS.Range(AWBS.Cells(1, 1), AWBS.Cells(AWBS_AC_LR, AWBS_LC_1R)), .Match(S4.Cells(FOR1, 1), AWBS.Range("A:A"), 0), 2)
''棚卸単価
'S4.Cells(FOR1, 4).Value = _
'.Index(AWBS.Range(AWBS.Cells(1, 1), AWBS.Cells(AWBS_AC_LR, AWBS_LC_1R)), .Match(S4.Cells(FOR1, 1), AWBS.Range("A:A"), 0), 13)
'
'Next FOR1
'
'AWB.Close savechanges:=False
'
'S4.Activate
'S4.Cells(S4_BC_LR + 1, 3).Activate
'Application.ScreenUpdating = True
'
'MsgBox "ブランド分類を入力してください"
'Exit Sub
'
'End If
'
'   End With
'
' 'ピボットで集計
'   With S7
'
'   .Activate
'
' .Cells.ClearContents
'
'  S1_BC_LR = S1.Cells(Rows.Count, 2).End(xlUp).Row
'
'    ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        S1.Range("A1:M" & S1_BC_LR)).CreatePivotTable _
'        TableDestination:=.Range("A2"), TableName:="ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1"
'
'End With
'
'    With S7.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("月")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'
'    With S7.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("（コード）")
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'
'    With S7
'
'    .PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").AddDataField .PivotTables( _
'        "ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("純売上数"), "合計 / 純売上数", xlSum
'    .PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").AddDataField .PivotTables( _
'        "ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("取引合計"), "合計 / 取引合計", xlSum
'    .PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("月").Subtotals = Array( _
'        False, False, False, False, False, False, False, False, False, False, False, False)
'
'        End With
'
'    With S7.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("月")
'
'        .LayoutForm = xlTabular
'
'    End With
'
'    With S7.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1")
'
'        .ColumnGrand = False
'        .RowGrand = False
'
'    End With
'
'S7.Cells.Copy
'
'   With S2
'
'    .AutoFilterMode = False
'
'.Activate
'.Cells.PasteSpecial Paste:=xlPasteValues
'.Rows("1:2").Delete Shift:=xlUp
'
' S2_BC_LR = .Cells(Rows.Count, 2).End(xlUp).Row
'
' '月DATEを入力
'
' For FOR1 = 3 To S2_BC_LR
'
' If .Cells(FOR1, 1).Value = "" Then
'
' .Cells(FOR1, 1).Value = .Cells(FOR1 - 1, 1).Value
'
' End If
'
' Next FOR1
'
'
'
'  .Cells(2, 5).Formula = "=INDEX(商品分類MST!$1:$1048576,MATCH($B2,商品分類MST!$A:$A,0),3)"
''  .Cells(2, 6).Formula = "=INDEX(商品分類MST!$1:$1048576,MATCH($B2,商品分類MST!$A:$A,0),4)"
''  .Cells(2, 7).Formula = "=INDEX(商品分類MST!$1:$1048576,MATCH($B2,商品分類MST!$A:$A,0),7)"
''    .Cells(2, 8).Formula = "=INDEX(商品分類MST!$1:$1048576,MATCH($B2,商品分類MST!$A:$A,0),8)"
''    .Cells(2, 9).Formula = "=INDEX(商品分類MST!$1:$1048576,MATCH($B2,商品分類MST!$A:$A,0),9)"
'    .Range(.Cells(2, 5), .Cells(2, 5)).Copy .Range(.Cells(3, 5), .Cells(S2_BC_LR, 5))
'
'    .Calculate
'
'' .Range("E:I").Copy
'' .Range("E:I").PasteSpecial Paste:=xlPasteValues
''
'' .Rows("1:1").AutoFilter Field:=7, Criteria1:="<>0"
''
''  .Cells.Copy
''  S3.Cells.PasteSpecial Paste:=xlPasteValues
''    .AutoFilterMode = False
''
' End With
'
'' With S3
'
'
' With S8
'
' .Activate
'
' Application.Calculation = xlCalculationAutomatic
'    Application.Calculation = xlCalculationManual
'
'
' If (Int(WorksheetFunction.Sum(S2.Range("D2:D" & S2_BC_LR))) * 1 _
' - Int(.Range("R96")) * 1 < 5) And (Int(WorksheetFunction.Sum(S2.Range("D2:D" & S2_BC_LR))) * 1 _
' - Int(.Range("R96")) * 1 > -5) Then
'
'
'  AWBN = ActiveWorkbook.Name
'
'  Application.DisplayAlerts = False
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\" & AWBN
' Application.DisplayAlerts = True
'
' Sheets(Array("部門別商品別", "部門別", "商品別(生活)", "商品別(服飾)", "商品別(バラエティ)")).Copy
'
'    Set AWS8 = ActiveWorkbook.Worksheets("部門別商品別")
'    Set AWS11 = ActiveWorkbook.Worksheets("部門別")
'    Set AWS12 = ActiveWorkbook.Worksheets("商品別(生活)")
'    Set AWS13 = ActiveWorkbook.Worksheets("商品別(服飾)")
'    Set AWS14 = ActiveWorkbook.Worksheets("商品別(バラエティ)")
''    Set AWS10 = ActiveWorkbook.Worksheets("シリーズ別(売上)")
'
'    AWS8.Cells.Copy
'    AWS8.Cells.PasteSpecial Paste:=xlPasteValues
''    AWS10.Cells.Copy
''    AWS10.Cells.PasteSpecial Paste:=xlPasteValues
'     AWS11.Cells.Copy
'    AWS11.Cells.PasteSpecial Paste:=xlPasteValues
'     AWS12.Cells.Copy
'    AWS12.Cells.PasteSpecial Paste:=xlPasteValues
'     AWS13.Cells.Copy
'    AWS13.Cells.PasteSpecial Paste:=xlPasteValues
'  AWS14.Cells.Copy
'    AWS14.Cells.PasteSpecial Paste:=xlPasteValues
'
' Application.DisplayAlerts = False
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\" & Month(Date) & "月" & Day(Date) & "日" & "部門別売上集計表.xlsx"
' Application.DisplayAlerts = True
'
' AWBN = Month(Date) & "月" & Day(Date) & "日" & "部門別売上集計表"
'
' Sheets(Array("部門別商品別", "部門別", "商品別(生活)", "商品別(服飾)", "商品別(バラエティ)")).Select
'    Sheets("部門別商品別").Activate
'
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'        "C:\Users\toyocase\Desktop\日報\" & AWBN & ".pdf", Quality:=xlQualityMinimum, _
'        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'        False
'
'        Set AWB = ActiveWorkbook
'
'        Application.DisplayAlerts = False
'        AWB.Worksheets("部門別商品別").Activate
'
'        AWB.Close savechanges:=False
'        Application.DisplayAlerts = True
'
' Else
'
' S4.Activate
' MsgBox "エラー:売上合計が合いません。分類見直し"
' Application.ScreenUpdating = True
'
'  Exit Sub
'
' End If
'
'  End With
'
'     Application.Calculation = xlCalculationAutomatic
'   Application.ScreenUpdating = True
'   MsgBox "正常終了"
'End Sub