Option Explicit

Public Sub import_csv_file(sheet_name As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheet_name)

    Dim last_row As Long
    last_row = get_last_row(ws, "A")
    ws.Rows("1:" & last_row).Delete

    ws.Cells.Style = "Normal"

    Dim query_table As QueryTable
    Dim file_open As Variant

    file_open = Application.GetOpenFilename(sheet_name & "CSVファイル (*.csv),*.csv", , sheet_name & "CSVファイルを選択してください")
    If file_open <> "False" Then     ' ファイルが選択された場合のみ処理を行う
        ' クエリテーブルを作成
        Set query_table = ws.QueryTables.Add(Connection:="TEXT;" & file_open, Destination:=ws.Range("$A$1"))
        With query_table
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
            .TextFileColumnDataTypes = Array(1, 2, 5, 1, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 2)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With

        ' クエリテーブルを削除
        For Each query_table In ws.QueryTables
            query_table.Delete
        Next query_table

        ' 接続を削除
        Dim wb_connection As WorkbookConnection
        For Each wb_connection In ThisWorkbook.Connections
            wb_connection.Delete
        Next wb_connection
    Else
        Exit Sub
    End If
End Sub

Public Sub toggle_execution_speed(ByVal enable As Boolean)
    Application.ScreenUpdating = Not enable
    Application.Calculation = IIf(enable, xlCalculationAutomatic, xlCalculationManual)
End Sub

Public Function get_last_row(ws As Worksheet, column As String) As Long
    Dim last_row As Long
    last_row = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
    get_last_row = last_row
End Function