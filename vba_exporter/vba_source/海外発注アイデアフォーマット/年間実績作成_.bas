Public Sub a1_年間実績作成()

    Dim ws As Worksheet
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    Dim last_row_ws1 As Long
    Dim last_row_ws2 As Long
    Dim last_row_ws3 As Long
    Dim last_row_ws4 As Long
    Dim i As Long
    Dim lookup_value As Variant
    Dim count_result As Long
    Dim search_range As Range

    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode = True Then
            ws.AutoFilterMode = False
        End If
    Next ws

    Set ws1 = ThisWorkbook.Worksheets("在庫表兼発注アイデア")

    Set ws2 = ThisWorkbook.Worksheets("年間売上数実績")

    Set ws3 = ThisWorkbook.Worksheets("集計")

    Set ws4 = ThisWorkbook.Worksheets("入力シート")

    Call import_csv_file("年間売上数実績")

    last_row_ws1 = get_last_row(ws1, "D")
    last_row_ws2 = get_last_row(ws2, "D")

    If last_row_ws2 < 2 Then
        Call toggle_execution_speed(True)
        Exit Sub
    End If

    Call toggle_execution_speed(False)

    Set search_range = ws1.Range("D3:D" & last_row_ws1)

    For i = 2 To last_row_ws2

        lookup_value = ws2.Cells(i, "D").Value

        If Not IsEmpty(lookup_value) And lookup_value <> "" Then

            count_result = Application.WorksheetFunction.CountIf(search_range, lookup_value)

            If count_result = 0 Then
                ws2.Cells(i, "K").Value = 1
            End If
        End If
    Next i
    ws2.Cells(last_row_ws2 + 1, "K").Value = 1

    last_row_ws2 = get_last_row(ws2, "K")
    If last_row_ws2 < 1 Then
        Call toggle_execution_speed(True)
        Exit Sub
    End If

    ws2.Range("A1:K" & last_row_ws2).AutoFilter Field:=11, Criteria1:="<>"
    ws2.Rows("2:" & last_row_ws2).Delete Shift:=xlUp

    If ws2.AutoFilterMode Then
        ws2.AutoFilterMode = False
    End If

    Call toggle_execution_speed(True)

    last_row_ws3 = get_last_row(ws3, "B")

    Debug.Print last_row_ws3

    ws3.Range("A1:W" & last_row_ws3).AutoFilter Field:=2, Criteria1:="<>""", _
    Operator:=xlAnd, _
    Criteria2:="<>0"

    last_row_ws4 = get_last_row(ws4, "B")

    ws4.Range("A1:S" & last_row_ws4).AutoFilter Field:=2, Criteria1:="<>""", _
    Operator:=xlAnd, _
    Criteria2:="<>0"

    Call toggle_execution_speed(False)

    Call create_custom_pivot_table("年間売上数実績", "年間実績表")

    Call toggle_execution_speed(True)
End Sub

Public Sub create_custom_pivot_table(source_sheet_name As String, pivot_sheet_name As String)

    Dim new_pivot_sheet As Worksheet
    Dim source_data_sheet As Worksheet
    Dim source_data_range As Range
    Dim last_row As Long
    Dim last_col As Long
    Dim pivot_cache As PivotCache
    Dim pivot_table As PivotTable
    Dim pt_field As PivotField    ' PivotFieldオブジェクト用
    Dim formula_start_row As Long
    Dim formula_end_row As Long
    Dim i As Long    ' ループカウンタ
    Dim formula_string As String

    formula_start_row = 5    ' 数式入力開始行

    ' --- 処理速度最適化 ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False    ' シート削除時の確認メッセージを非表示

    ' --- データソースシートの存在確認 ---
    On Error Resume Next
    Set source_data_sheet = ThisWorkbook.Worksheets(source_sheet_name)
    On Error GoTo 0
    If source_data_sheet Is Nothing Then
        MsgBox "データソースシート '" & source_sheet_name & "' が見つかりません。", vbCritical
        GoTo ExitHandler    ' エラー時は後処理へ
    End If

    ' --- 既存の指定ピボットシートがあれば削除 ---
    On Error Resume Next
    Set new_pivot_sheet = ThisWorkbook.Worksheets(pivot_sheet_name)
    On Error GoTo 0
    If Not new_pivot_sheet Is Nothing Then
        new_pivot_sheet.Delete
    End If
    Set new_pivot_sheet = Nothing

    ' --- 新しいシートを作成し、指定した名前に変更 ---
    Dim ref_sheet As Worksheet    ' 基準となるシート（年間売上数実績）

    ' 基準シートオブジェクトを取得
    On Error Resume Next
    Set ref_sheet = ThisWorkbook.Worksheets(source_sheet_name)    ' "年間売上数実績"シート
    On Error GoTo 0

    If ref_sheet Is Nothing Then
        MsgBox "基準となるシート '" & source_sheet_name & "' が見つかりません。シートを追加できません。", vbCritical
        GoTo ExitHandler
    End If

    ' 基準シートの左隣に新しいシートを追加
    Set new_pivot_sheet = ThisWorkbook.Worksheets.Add(Before:=ref_sheet)

    ' 新しいシートに名前を設定
    On Error Resume Next
    new_pivot_sheet.Name = pivot_sheet_name
    If Err.Number <> 0 Then
        MsgBox "シート名 '" & pivot_sheet_name & "' の設定に失敗しました。" & vbCrLf & _
        "既に同名のシートが存在する可能性があります。", vbCritical
        Err.Clear
        ' 必要であれば、ここで追加したシートを削除するなどの後処理を追加
        ' new_pivot_sheet.Delete ' など
        GoTo ExitHandler
    End If
    On Error GoTo 0

    ' --- ピボットテーブルのデータソース範囲を動的に取得 ---
    With source_data_sheet
        If Application.WorksheetFunction.CountA(.Cells) = 0 Then
            MsgBox "データソースシート '" & source_sheet_name & "' にデータがありません。", vbExclamation
            GoTo ExitHandler
        End If
        last_row = .Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        last_col = .Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).column
        If last_row = 0 Or last_col = 0 Then
            MsgBox "データソースシート '" & source_sheet_name & "' のデータ範囲を特定できませんでした。", vbExclamation
            GoTo ExitHandler
        End If
        Set source_data_range = .Range(.Cells(1, 1), .Cells(last_row, last_col))
    End With

    ' --- ピボットキャッシュの作成 ---
    On Error Resume Next
    Set pivot_cache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
    SourceData:=source_data_range, _
    Version:=xlPivotTableVersion12)
    If Err.Number <> 0 Then
        MsgBox "ピボットキャッシュの作成に失敗しました。" & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitHandler
    End If
    On Error GoTo 0

    ' --- ピボットテーブルの作成 ---
    On Error Resume Next
    Set pivot_table = pivot_cache.CreatePivotTable(TableDestination:=new_pivot_sheet.Range("A3"), _
    TableName:="ピボットテーブル1", _
    DefaultVersion:=xlPivotTableVersion12)
    If Err.Number <> 0 Then
        MsgBox "ピボットテーブルの作成に失敗しました。" & vbCrLf & Err.Description, vbCritical
        Err.Clear
        GoTo ExitHandler
    End If
    On Error GoTo 0

    ' --- ピボットテーブルの書式設定とフィールド設定 ---
    With pivot_table
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .CompactRowIndent = 1
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .RowAxisLayout xlCompactRow
    End With

    With pivot_table.PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With

    On Error Resume Next
    pivot_table.RepeatAllLabels xlRepeatLabels
    On Error GoTo 0

    ' --- フィールド設定 ---
    Dim field_name As String
    field_name = "担当者名"
    On Error Resume Next
    Set pt_field = pivot_table.PivotFields(field_name)
    If Err.Number = 0 Then
        pt_field.Orientation = xlPageField
        pt_field.Position = 1
    Else
        Debug.Print "フィルターフィールド '" & field_name & "' が見つかりません。"
        Err.Clear
    End If
    On Error GoTo 0
    Set pt_field = Nothing

    field_name = "実績年月"
    On Error Resume Next
    Set pt_field = pivot_table.PivotFields(field_name)
    If Err.Number = 0 Then
        pt_field.Orientation = xlColumnField
        pt_field.Position = 1
    Else
        Debug.Print "フィールド '" & field_name & "' が見つかりません。"
        Err.Clear
    End If
    On Error GoTo 0
    Set pt_field = Nothing

    field_name = "商品コード"
    On Error Resume Next
    Set pt_field = pivot_table.PivotFields(field_name)
    If Err.Number = 0 Then
        pt_field.Orientation = xlRowField
        pt_field.Position = 1
    Else
        Debug.Print "フィールド '" & field_name & "' が見つかりません。"
        Err.Clear
    End If
    On Error GoTo 0
    Set pt_field = Nothing

    field_name = "得意先略称"
    On Error Resume Next
    Set pt_field = pivot_table.PivotFields(field_name)
    If Err.Number = 0 Then
        pt_field.Orientation = xlRowField
        pt_field.Position = 2
    Else
        Debug.Print "フィールド '" & field_name & "' が見つかりません。"
        Err.Clear
    End If
    On Error GoTo 0
    Set pt_field = Nothing

    field_name = "売上数量"
    On Error Resume Next
    Set pt_field = pivot_table.PivotFields(field_name)
    If Err.Number = 0 Then
        pivot_table.AddDataField pt_field, "合計 / " & field_name, xlSum
    Else
        Debug.Print "データフィールド '" & field_name & "' の追加に失敗しました。"
        Err.Clear
    End If
    On Error GoTo 0
    Set pt_field = Nothing

    ' --- 後処理 ---
    ExitHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

    If Err.Number = 0 And Not pivot_table Is Nothing Then
        ' ウィンドウ枠の固定とフィールドリスト非表示
        new_pivot_sheet.Activate
        new_pivot_sheet.Range("B5").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        ActiveWorkbook.ShowPivotTableFieldList = False

        ' ★★★ ここから数式入力処理 ★★★
        ' ピボットテーブルのデータ本体の最終行を取得 (総計行を除く)
        ' TableRange2 はピボットテーブル全体（ページフィールド含む）を返す
        ' DataBodyRange はデータ本体（ヘッダと総計を除く）を返す
        On Error Resume Next    ' DataBodyRangeが取得できないケース（データがないなど）を考慮
        If pivot_table.DataBodyRange Is Nothing Then
            Debug.Print "ピボットテーブルにデータがありません。数式は入力されません。"
        Else
            ' DataBodyRange の最終行を基準にするが、ピボットテーブルはA3から始まっているのでオフセットを考慮
            ' formula_end_row = pivot_table.TableRange1.Rows.Count + pivot_table.TableRange1.Row - 1 ' TableRange1はヘッダ含む
            ' より確実なのは、ピボットテーブルの行フィールドのアイテム数を数えるか、
            ' DataBodyRange の最終行を使う
            ' ピボットテーブルはA3から開始。行フィールドのアイテムがA列に表示されると仮定。
            ' A列のデータがある最終行をピボットの最終行とするのは、レイアウトに依存する
            ' ここでは、ピボットテーブルの RowRange (行フィールドとデータラベルの範囲) を使うことを試みる
            ' ただし、総計行を含む場合があるので注意

            ' ピボットテーブルの行フィールドのデータがある最終行を取得する試み
            ' (A列に商品コードが表示され、それが最初の行フィールドであると仮定)
            Dim last_pivot_row As Long
            ' A列の最終データ行を取得 (ピボットテーブルの範囲内で)
            ' TableRange1 はピボットテーブル全体の範囲 (ページフィールドは含まないことが多い)
            If pivot_table.RowFields.count > 0 Then
                ' A列 (ピボットテーブルの開始列) の最終行
                last_pivot_row = new_pivot_sheet.Cells(Rows.count, pivot_table.TableRange1.column).End(xlUp).Row
                ' ただし、総計行が含まれている場合はその行を除く
                If pivot_table.RowGrand Then
                    If new_pivot_sheet.Cells(last_pivot_row, pivot_table.TableRange1.column).Value = "総計" Or _
                        LCase(new_pivot_sheet.Cells(last_pivot_row, pivot_table.TableRange1.column).Value) = "grand total" Then
                        last_pivot_row = last_pivot_row - 1
                    End If
                End If
                formula_end_row = last_pivot_row
            Else
                formula_end_row = formula_start_row    ' 行フィールドがない場合はP5のみ
            End If

            If formula_end_row >= formula_start_row Then
                ' 数式文字列 (ダブルクォーテーションは二重にする)
                ' 数式内のシート名は正確に指定してください
                formula_string = "=IF(COUNTIF(在庫表兼発注アイデア!$D:$D,$A5)>0,IF($B$1=""(すべて)"",SUMIFS(年間売上数実績!$J:$J,年間売上数実績!$D:$D,$A5),SUMIFS(年間売上数実績!$J:$J,年間売上数実績!$D:$D,$A5,年間売上数実績!$G:$G,$B$1)),"""")"

                ' P列に数式を入力
                ' new_pivot_sheet.Range("P" & formula_start_row & ":P" & formula_end_row).Formula = formula_string
                ' 上記だと、$A5の部分がフィルダウンされない。R1C1形式で対応するか、1行ずつ入力。

                ' 1行ずつ数式を入力して、相対参照を正しく機能させる
                For i = formula_start_row To formula_end_row
                    ' 各行に合わせて $A5 の部分を $A & i に変更
                    Dim current_formula As String
                    current_formula = "=IF(COUNTIF(在庫表兼発注アイデア!$D:$D,$A" & i & ")>0,IF($B$1=""(すべて)"",SUMIFS(年間売上数実績!$J:$J,年間売上数実績!$D:$D,$A" & i & "),SUMIFS(年間売上数実績!$J:$J,年間売上数実績!$D:$D,$A" & i & ",年間売上数実績!$G:$G,$B$1)),"""")"
                    new_pivot_sheet.Cells(i, "P").Formula = current_formula
                Next i
                new_pivot_sheet.Cells(4, "P").Value = "売上金額"
                new_pivot_sheet.Columns("P:P").NumberFormatLocal = "#,##0_ ;[赤]-#,##0 "

                Debug.Print "P列に数式を入力しました。範囲: P" & formula_start_row & " から P" & formula_end_row
            Else
                Debug.Print "数式を入力する有効な行範囲がありません。"
            End If
        End If
        On Error GoTo 0
        ' ★★★ 数式入力処理ここまで ★★★

        MsgBox "ピボットテーブルがシート '" & pivot_sheet_name & "' に作成されました。", vbInformation
    ElseIf Err.Number <> 0 Then
        ' エラー発生時は既にメッセージ表示済み
    End If

    ' オブジェクト解放
    Set pt_field = Nothing
    Set pivot_table = Nothing
    Set pivot_cache = Nothing
    Set source_data_range = Nothing
    Set new_pivot_sheet = Nothing
    Set source_data_sheet = Nothing

End Sub