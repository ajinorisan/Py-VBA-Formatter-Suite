Option Explicit

Sub データ取込_1()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("貼付")
    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("出力")
    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("商品マスタ")
    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("拠点マスタ")
    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("返品データ加工")
    Dim WS6 As Worksheet
    Set WS6 = ThisWorkbook.Worksheets("補填データ加工")
    Dim WS7 As Worksheet
    Set WS7 = ThisWorkbook.Worksheets("売上取込フォーマット")

    WS1.Cells.ClearContents
    WS6.Columns("A:L").Delete
    WS6.Cells.ClearContents

    ' ファイル選択ダイアログを表示
    Dim FILE_DIALOG As Object
    Set FILE_DIALOG = Application.FileDialog(msoFileDialogFilePicker)
    Dim FIRST As Boolean
    FIRST = True

    With FILE_DIALOG
        .Title = "藤栄返品CSVファイルを全て選択してください"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"

        If .Show = -1 Then    ' OKボタンが押された場合
            ' 選択された全てのファイルを取り込む
            Dim FILE_TO_OPEN As Variant
            For Each FILE_TO_OPEN In .SelectedItems
                ' クエリテーブルを作成
                If FIRST Then
                    With WS1.QueryTables.Add(Connection:="TEXT;" & FILE_TO_OPEN, Destination:=WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Offset(0, 0))

                        .FieldNames = True
                        .RowNumbers = False
                        .FillAdjacentFormulas = False
                        .PreserveFormatting = True
                        .RefreshOnFileOpen = False
                        .RefreshStyle = xlInsertDeleteCells
                        .SavePassword = False
                        .SaveData = True
                        .AdjustColumnWidth = True
                        .RefreshPeriod = 0
                        .TextFilePromptOnRefresh = False
                        .TextFilePlatform = 932
                        .TextFileStartRow = 1
                        .TextFileParseType = xlDelimited
                        .TextFileTextQualifier = xlTextQualifierDoubleQuote
                        .TextFileConsecutiveDelimiter = False
                        .TextFileTabDelimiter = False
                        .TextFileSemicolonDelimiter = False
                        .TextFileCommaDelimiter = True
                        .TextFileSpaceDelimiter = False
                        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 2, 2)
                        .TextFileTrailingMinusNumbers = True
                        .Refresh BackgroundQuery:=False
                        FIRST = False
                    End With
                Else
                    With WS1.QueryTables.Add(Connection:="TEXT;" & FILE_TO_OPEN, Destination:=WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Offset(1, 0))

                        .FieldNames = True
                        .RowNumbers = False
                        .FillAdjacentFormulas = False
                        .PreserveFormatting = True
                        .RefreshOnFileOpen = False
                        .RefreshStyle = xlInsertDeleteCells
                        .SavePassword = False
                        .SaveData = True
                        .AdjustColumnWidth = True
                        .RefreshPeriod = 0
                        .TextFilePromptOnRefresh = False
                        .TextFilePlatform = 932
                        .TextFileStartRow = 1
                        .TextFileParseType = xlDelimited
                        .TextFileTextQualifier = xlTextQualifierDoubleQuote
                        .TextFileConsecutiveDelimiter = False
                        .TextFileTabDelimiter = False
                        .TextFileSemicolonDelimiter = False
                        .TextFileCommaDelimiter = True
                        .TextFileSpaceDelimiter = False
                        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 2, 2)
                        .TextFileTrailingMinusNumbers = True
                        .Refresh BackgroundQuery:=False
                    End With
                End If
            Next FILE_TO_OPEN

            Dim QT As Variant
            ' クエリテーブルを削除
            For Each QT In WS1.QueryTables
                QT.Delete
            Next QT

            ' 接続を削除
            Dim CN As WorkbookConnection
            For Each CN In ThisWorkbook.Connections
                CN.Delete
            Next CN
        Else
            Exit Sub    ' キャンセルが押された場合は終了
        End If
    End With

    FIRST = True

    Dim WB As Workbook
    Dim WS As Worksheet
    Dim LAST_ROW As Long

    With FILE_DIALOG
        .Title = "藤栄値引XLSファイルを全て選択してください"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "XLSファイル", "*.xls"

        If .Show = -1 Then    ' OKボタンが押された場合
            ' 選択された全てのファイルを取り込む
            'Dim FILE_TO_OPEN As Variant
            For Each FILE_TO_OPEN In .SelectedItems
                ' クエリテーブルを作成
                If FIRST Then
                    Set WB = Workbooks.Open(FILE_TO_OPEN)
                    Set WS = ActiveWorkbook.Worksheets("Sheet1")

                    With WS
                        LAST_ROW = .Cells(.Rows.Count, 1).End(xlUp).Row
                        If LAST_ROW >= 3 Then
                            .Range("A3").Resize(LAST_ROW - 2, .Columns.Count).Copy WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Offset(0, 0)
                        End If
                    End With

                    WB.Close SaveChanges:=False
                    FIRST = False
                Else
                    Set WB = Workbooks.Open(FILE_TO_OPEN)
                    Set WS = ActiveWorkbook.Worksheets("Sheet1")

                    With WS
                        LAST_ROW = .Cells(.Rows.Count, 1).End(xlUp).Row
                        If LAST_ROW >= 4 Then
                            .Range("A4").Resize(LAST_ROW - 3, .Columns.Count).Copy WS6.Cells(WS6.Rows.Count, 1).End(xlUp).Offset(1, 0)
                        End If
                    End With
                    WB.Close SaveChanges:=False
                End If
            Next FILE_TO_OPEN

            ' ... (後続のコードと同様)

        Else
            Exit Sub    ' キャンセルが押された場合は終了
        End If
    End With

    Dim WS1_AC_LR As Long
    WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

    WS5.Cells.ClearContents
    WS1.Range("A1:R1").Copy WS5.Range("A1")
    Dim WS5_AC_LR As Long
    WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row

    Dim i As Long

    For i = 2 To WS1_AC_LR

        If WS1.Cells(i, 2) <> "支店名" And WS1.Cells(i, 6) = "返品" Then

            WS1.Range(WS1.Cells(i, 1), WS1.Cells(i, 18)).Copy WS5.Cells(WS5_AC_LR + 1, 1)
            WS5_AC_LR = WS5.Cells(WS5.Rows.Count, 1).End(xlUp).Row
        End If

    Next i

    WS5.Activate
    MsgBox "A列支店コード修正と計上済のデータの削除"
End Sub