Option Explicit



'Private Sub CommandButton1_Click()
'
'Dim 個人別集計 As String
'個人別集計 = "C:\Users\toyocase\Desktop\日報用データ\売上集計_担当別_個人売上用.csv"
'
'
'Dim S1 As Worksheet
'Const Sheet1 As String = "個人売上"
'Set S1 = ThisWorkbook.Worksheets(Sheet1)
'Dim S2 As Worksheet
'Const Sheet2 As String = "売上集計表"
'Set S2 = ThisWorkbook.Worksheets(Sheet2)
'Dim S3 As Worksheet
'Const Sheet3 As String = "ピボット"
'Set S3 = ThisWorkbook.Worksheets(Sheet3)
'Dim S4 As Worksheet
'Const Sheet4 As String = "58期"
'Set S4 = ThisWorkbook.Worksheets(Sheet4)
'Dim S5 As Worksheet
'Const Sheet5 As String = "営業リスト"
'Set S5 = ThisWorkbook.Worksheets(Sheet5)
'Dim S5_A列_LR As Integer
'Dim S4_A列_LR As Integer
'
' Application.ScreenUpdating = False
'
'
' Workbooks.Open Filename:= _
'        個人別集計, ReadOnly:=True
'        Dim AWB As Workbook
'        Set AWB = ActiveWorkbook
'        AWB.ActiveSheet.Cells.Select
'        Selection.Copy
'        S2.Activate
'
'        Cells.PasteSpecial Paste:=xlPasteValues
'       Application.CutCopyMode = False
'AWB.Close savechanges:=False
'
'With S2
'
'Dim S2_売上数列_LR As Long
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'Dim FOR1 As Long
'If .Range("B1").Value = "売上集計表" Then
'.Rows("1:1").Delete
'
''オートフィルタかけ直し
'If .AutoFilterMode = True Then
'.AutoFilterMode = False
'
''不必要データ削除
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=2, Criteria1:=Array( _
'        "事業部<<合計>>", "総合計", "日付<<合計>>"), Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'Else
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=2, Criteria1:=Array( _
'        "事業部<<合計>>", "総合計", "日付<<合計>>"), Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'End If
'End If
'
''行削除のため再度、売上集計表シートの売上数列の最終行取得
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'
''完成データかを確認
'If .Cells(1, 1) = "" Then
'
'.Cells(1, 1).Value = "売上月"
'
'For FOR1 = 2 To S2_売上数列_LR
'If .Cells(FOR1, 3).Value = "" Then
'.Cells(FOR1, 1).Value = _
'Year(.Cells(FOR1, 2)) & Right("0" & Month(.Cells(FOR1, 2)), 2)
'Else
'.Cells(FOR1, 1).Value = _
'.Cells(FOR1 - 1, 1).Value
'End If
'Next FOR1
'
'If .AutoFilterMode = True Then
'.AutoFilterMode = False
'
''不必要データ削除
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=3, Criteria1:="", Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'Else
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=3, Criteria1:="", Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'End If
'End If
'
''行削除のため再度、売上集計表シートの売上数列の最終行取得
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'
''集計表データをピボットで月ごとに集計
'S3.Activate
'S3.Cells.ClearContents
'
'ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        .Range("A1:M" & S2_売上数列_LR)).CreatePivotTable TableDestination:=S3.Range("A2"), TableName:="ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1"
'
'With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("売上月")
'        .Orientation = xlColumnField
'        .Position = 1
'    End With
'
'    With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("（コード）")
'    .LayoutCompactRow = False
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'
'      With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("名称")
'
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'
'    S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").AddDataField S3.PivotTables( _
'        "ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("取引合計"), "合計 / 取引合計", xlSum
'
'S3.Cells.Copy
'S4.Activate
'S4.Cells.PasteSpecial Paste:=xlPasteValues
'S5_A列_LR = S5.Cells(Rows.Count, 1).End(xlUp).Row
'S4_A列_LR = S4.Cells(Rows.Count, 1).End(xlUp).Row
'
'For FOR1 = 4 To S4_A列_LR - 2 Step 2
'If WorksheetFunction.CountIf(S5.Range("A1:A" & S5_A列_LR), S4.Cells(FOR1, 1)) = 0 Then
'MsgBox S4.Cells(FOR1, 1).Value & " " & S4.Cells(FOR1 + 1, 2).Value & " " & "営業担当を修正"
'Application.ScreenUpdating = True
'Exit Sub
'End If
'Next FOR1
' Application.Calculation = xlAutomatic
''Const 日報場所 As String = "\\192.168.2.2\share\竜王共有ファイル\日報"
'Dim 日報場所 As String
'日報場所 = "C:\Users\toyocase\Desktop\日報"
'Dim DATE1 As Date
'DATE1 = Date
'
' Dim AWBN As String
'  AWBN = ActiveWorkbook.Name
'  S1.Activate
'  Application.DisplayAlerts = False
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\個人売上日報.xlsm"
' Application.DisplayAlerts = True
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\" & Month(DATE1) & "月" & Day(DATE1) & "日" & AWBN
'
' AWBN = Month(DATE1) & "月" & Day(DATE1) & "日" & "個人別売上日報"
'
'S1.Activate
'    S1.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'        "C:\Users\toyocase\Desktop\日報\" & AWBN & ".pdf", Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'        False
'
'S1.Activate
'
'Dim GOSデータ場所 As String
'GOSデータ場所 = "\\192.168.2.2\share\竜王共有ファイル\GOSマスタ(名称･場所変更禁止)"
'Dim 日報用データ場所 As String
'日報用データ場所 = "C:\Users\toyocase\Desktop\日報用データ"
'
'Dim fso
'Dim strPath1, strPath2, strPath3, strPath4
'
'strPath1 = 日報用データ場所 & "\商品一覧.csv" 'コピー元
'strPath2 = GOSデータ場所 & "\商品一覧.csv"  'コピー先
'
'strPath3 = 日報用データ場所 & "\得意先一覧.csv" 'コピー元
'strPath4 = GOSデータ場所 & "\得意先一覧.csv"  'コピー先
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'
'fso.CopyFile strPath1, strPath2
'fso.CopyFile strPath3, strPath4
'
'Set fso = Nothing
'
' Application.ScreenUpdating = True
' MsgBox "更新終了"
'End With
'
'End Sub

'Private Sub CommandButton21_Click() '得意先別売上対比表作成
'
'
'Dim S1 As Worksheet
'Const sheet1 As String = "個人売上"
'Set S1 = ThisWorkbook.Worksheets(sheet1)
'Dim S2 As Worksheet
'Const sheet2 As String = "売上集計表"
'Set S2 = ThisWorkbook.Worksheets(sheet2)
'Dim S3 As Worksheet
'Const Sheet3 As String = "ピボット"
'Set S3 = ThisWorkbook.Worksheets(Sheet3)
'Dim S4 As Worksheet
'Const Sheet4 As String = "57期"
'Set S4 = ThisWorkbook.Worksheets(Sheet4)
'Dim S5 As Worksheet
'Const Sheet5 As String = "営業リスト"
'Set S5 = ThisWorkbook.Worksheets(Sheet5)
'Dim S5_A列_LR As Integer
'Dim S4_A列_LR As Integer
'
' Application.ScreenUpdating = False
'
'With S2
'
'Dim S2_売上数列_LR As Long
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'Dim FOR1 As Long
'If .Range("B1").Value = "売上集計表" Then
'.Rows("1:1").Delete
'
''オートフィルタかけ直し
'If .AutoFilterMode = True Then
'.AutoFilterMode = False
'
''不必要データ削除
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=2, Criteria1:=Array( _
'        "事業部<<合計>>", "総合計", "日付<<合計>>"), Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'Else
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=2, Criteria1:=Array( _
'        "事業部<<合計>>", "総合計", "日付<<合計>>"), Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'End If
'End If
'
''行削除のため再度、売上集計表シートの売上数列の最終行取得
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'
''完成データかを確認
'If .Cells(1, 1) = "" Then
'
'.Cells(1, 1).Value = "売上月"
'
'For FOR1 = 2 To S2_売上数列_LR
'If .Cells(FOR1, 3).Value = "" Then
'.Cells(FOR1, 1).Value = _
'Year(.Cells(FOR1, 2)) & Right("0" & Month(.Cells(FOR1, 2)), 2)
'Else
'.Cells(FOR1, 1).Value = _
'.Cells(FOR1 - 1, 1).Value
'End If
'Next FOR1
'
'If .AutoFilterMode = True Then
'.AutoFilterMode = False
'
''不必要データ削除
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=3, Criteria1:="", Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'Else
'.Range(.Cells(1, 1), .Cells(S2_売上数列_LR, 13)).AutoFilter Field:=3, Criteria1:="", Operator:=xlFilterValues
'        .Rows("2:" & S2_売上数列_LR).Delete shift:=xlUp
'        .AutoFilterMode = False
'End If
'End If
'
''行削除のため再度、売上集計表シートの売上数列の最終行取得
'S2_売上数列_LR = .Cells(Rows.Count, 4).End(xlUp).Row
'
''集計表データをピボットで月ごとに集計
'S3.Activate
'S3.Cells.ClearContents
'
'ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        .Range("A1:M" & S2_売上数列_LR)).CreatePivotTable TableDestination:=S3.Range("A2"), TableName:="ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1"
'
'With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("売上月")
'        .Orientation = xlColumnField
'        .Position = 1
'    End With
'
'    With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("（コード）")
'    .LayoutCompactRow = False
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'
'      With S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("名称")
'
'        .Orientation = xlRowField
'        .Position = 2
'    End With
'
'    S3.PivotTables("ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").AddDataField S3.PivotTables( _
'        "ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1").PivotFields("取引合計"), "合計 / 取引合計", xlSum
'
'S3.Cells.Copy
'S4.Activate
'S4.Cells.PasteSpecial Paste:=xlPasteValues
'S5_A列_LR = S5.Cells(Rows.Count, 1).End(xlUp).Row
'S4_A列_LR = S4.Cells(Rows.Count, 1).End(xlUp).Row
'
'For FOR1 = 4 To S4_A列_LR - 2 Step 2
'If WorksheetFunction.CountIf(S5.Range("A1:A" & S5_A列_LR), S4.Cells(FOR1, 1)) = 0 Then
'MsgBox S4.Cells(FOR1, 1).Value & " " & S4.Cells(FOR1 + 1, 2).Value & " " & "営業担当を修正"
'Application.ScreenUpdating = True
'Exit Sub
'End If
'Next FOR1
' Application.Calculation = xlAutomatic
'Const 日報場所 As String = "\\192.168.2.2\share\竜王共有ファイル\日報"
'Dim DATE1 As Date
'DATE1 = Date
'
'Dim AWBN As String
'  AWBN = ActiveWorkbook.Name
'   S1.Activate
'  Application.DisplayAlerts = False
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\個人売上日報.xlsm"
' Application.DisplayAlerts = True
' ActiveWorkbook.SaveAs Filename:=日報場所 & "\" & Month(DATE1) & "月" & Day(DATE1) & "日" & AWBN
'
' AWBN = Month(DATE1) & "月" & Day(DATE1) & "日" & "個人別売上日報"
'
'S1.Activate
'    S1.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'        "\\192.168.2.2\share\竜王共有ファイル\日報\" & AWBN & ".pdf", Quality:=xlQualityStandard, _
'        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
'        False
'
'S1.Activate
' Application.ScreenUpdating = True
' MsgBox "更新終了"
'End With
'End Sub