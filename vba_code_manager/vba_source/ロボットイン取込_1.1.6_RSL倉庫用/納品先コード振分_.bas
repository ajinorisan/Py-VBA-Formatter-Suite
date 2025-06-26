Option Explicit

Sub 納品先コード振分()

'売上日付指定、納品先コード振分、AO取り込み要データ作成

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("データ抽出用")


    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("アラジン取込用(売上)")

    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("納品先マスタ")

    Dim WS6 As Worksheet
    Set WS6 = ThisWorkbook.Worksheets("商品変換マスタ")

    Dim WS9 As Worksheet
    Set WS9 = ThisWorkbook.Worksheets("モールマスタ")

    Dim WS10 As Worksheet
    Set WS10 = ThisWorkbook.Worksheets("在庫用")

    Dim WS11 As Worksheet
    Set WS11 = ThisWorkbook.Worksheets("配送便確認用")


    Dim FOR1 As Long

    With WorksheetFunction

        If WS4.AutoFilterMode = True Then
            WS4.AutoFilterMode = False
        End If

        Dim WS4_EC_LR As Long
        WS4_EC_LR = WS4.Cells(WS4.Rows.Count, 5).End(xlUp).Row

        For FOR1 = 2 To WS4_EC_LR
            If WS4.Cells(FOR1, 10) <> WS4.Cells(FOR1 - 1, 10) Then
                WS4.Cells(FOR1, 29).Value = .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 39)

                'WS4.Cells(FOR1, 7).Value = .Index(WS5.Cells, .Match(納品先CD, WS5.Range("C:C"), 0), 1)
                'WS4.Cells(FOR1, 8).Formula = Mid(WS4.Cells(FOR1, 25), 1, 10) & " " & Format(WS4.Cells(FOR1, 16), "000")
                'WS4.Cells(FOR1, 28).Value = WS4.Cells(FOR1, 25) & .VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 11, False) _
                 '& .VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 12, False) & _
                 '.VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 13, False)
                '納品先CD = 納品先CD + 1
            Else
                WS4.Cells(FOR1, 29).Value = WS4.Cells(FOR1 - 1, 29)
                'WS4.Cells(FOR1, 8).Value = WS4.Cells(FOR1 - 1, 8)
            End If
        Next FOR1
    End With

    'モール毎並び替え
    '   WS4.Sort.SortFields.Clear
    '    WS4.Sort.SortFields.Add2 Key:=WS4.Range("E2:E" & WS4_EC_LR) _
         '    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    '    With WS4.Sort
    '        .SetRange WS4.Range("A1:AC" & WS4_EC_LR)
    '        .Header = xlYes
    '        .MatchCase = False
    '        .Orientation = xlTopToBottom
    '        .SortMethod = xlPinYin
    '        .Apply
    '    End With

    '売上データ日付セット
売上日付入力:
    Dim IB2 As String
    IB2 = Application.InputBox(Prompt:="売上日付入力(半角数字8桁/YYYYMMDD)" & vbCrLf & _
                                       "例)20220201", Type:=2)
    If IB2 = "" Then GoTo 売上日付入力
    If IB2 = "False" Then Exit Sub
    WS4.Range(WS4.Cells(2, 1), WS4.Cells(WS4_EC_LR, 2)).Value = IB2
    If Len(WS4.Range("A2")) <> 8 Then GoTo 売上日付入力

    '希望納期入力:
    'Dim IB3 As String
    'IB3 = Application.InputBox(Prompt:="希望納期入力(半角数字8桁/YYYYMMDD)" & vbCrLf & _
     '"例)20200526", Type:=2)
    'If IB3 = "" Then GoTo 希望納期入力
    'If IB3 = "False" Then Exit Sub
    'WS4.Range(WS4.Cells(2, 24), WS4.Cells(WS4_EC_LR, 24)).Value = IB3
    'If Len(WS4.Range("X2")) <> 8 Then GoTo 希望納期入力


納品先CD入力:
    'On Error GoTo 納品先CD入力
    Dim IB4 As String
    IB4 = Application.InputBox(Prompt:="納品先CD個人No?", Type:=2)
    If IB4 = "" Then GoTo 納品先CD入力
    If IB4 = "False" Then Exit Sub
    WS4.Range("P2").Value = IB4
    If Len(WS4.Range("P2")) >= 4 Then GoTo 納品先CD入力
    'On Error GoTo 0

    Dim 納品先CD As Long
    納品先CD = IB4

    With WorksheetFunction

        For FOR1 = 2 To WS4_EC_LR
            If WS4.Cells(FOR1, 10) <> WS4.Cells(FOR1 - 1, 10) Then
                WS4.Cells(FOR1, 16).Value = 納品先CD
                WS4.Cells(FOR1, 7).Value = .Index(WS5.Cells, .Match(納品先CD, WS5.Range("C:C"), 0), 1)
                WS4.Cells(FOR1, 8).Formula = _
                Mid(WS4.Cells(FOR1, 29), 1, 1) & " " & Mid(WS4.Cells(FOR1, 25), 1, 10) & " " & Format(WS4.Cells(FOR1, 16), "000")
                'WS4.Cells(FOR1, 28).Value = WS4.Cells(FOR1, 25) & .VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 11, False) _
                 & .VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 12, False) & _
                 .VLookup(WS4.Cells(FOR1, 25), WS2.Range("Q:AC"), 13, False)
                納品先CD = 納品先CD + 1
            Else
                WS4.Cells(FOR1, 7).Value = WS4.Cells(FOR1 - 1, 7)
                WS4.Cells(FOR1, 8).Value = WS4.Cells(FOR1 - 1, 8)
            End If

            If WS4.Cells(FOR1, 24) = 0 Then
                WS4.Cells(FOR1, 24).Value = WS4.Cells(FOR1, 1)
            End If

        Next FOR1

        ''同梱処理
        'For FOR1 = 2 To WS4_EC_LR
        'If .CountIfs(WS4.Range(WS4.Cells(2, 28), WS4.Cells(FOR1, 28)), WS4.Cells(FOR1, 28)) >= 2 Then
        'If WS4.Cells(FOR1 + 1, 28) = "" Then
        'WS4.Cells(FOR1, 28).Copy WS4.Cells(FOR1 + 1, 28)
        'End If
        'WS4.Cells(FOR1, 7).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 7)
        ''20230206アラジン仕様変更による同梱処理の変更
        'WS4.Cells(FOR1, 8).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 8)
        'WS4.Cells(FOR1, 16).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 16)
        'WS4.Cells(FOR1, 27).Value = 1
        'End If
        'If (WS4.Cells(FOR1 - 1, 27) = 1 And WS4.Cells(FOR1, 28) = "") Then
        'WS4.Cells(FOR1, 7).Value = WS4.Cells(FOR1 - 1, 7)
        'WS4.Cells(FOR1, 8).Value = WS4.Cells(FOR1 - 1, 8)
        '
        'End If
        'Next FOR1
        '
        'For FOR1 = 2 To WS4_EC_LR
        'If .CountIfs(WS4.Range(WS4.Cells(2, 28), WS4.Cells(FOR1, 28)), WS4.Cells(FOR1, 28)) >= 3 Then
        'WS4.Cells(FOR1, 27).ClearContents
        'End If
        'Next FOR1
        ''同梱処理終了


        '同梱処理
        For FOR1 = 2 To WS4_EC_LR
            If .CountIfs(WS4.Range(WS4.Cells(2, 28), WS4.Cells(FOR1, 28)), WS4.Cells(FOR1, 28)) >= 2 Then
On Error Resume Next
                WS4.Cells(FOR1, 7).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 7)
                
                '20230206アラジン仕様変更による同梱処理の変更
                WS4.Cells(FOR1, 8).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 8)
                WS4.Cells(FOR1, 16).Value = .Index(WS4.Cells, .Match(WS4.Cells(FOR1, 28), WS4.Range("AB：AB"), 0), 16)
                On Error GoTo 0
            End If

        Next FOR1

        '同一人物かつ基データが並んでいる場合同梱処理されないのを修正20230329
        For FOR1 = WS4_EC_LR To 2 Step -1
            If (WS4.Cells(FOR1, 16) = WS4.Cells(FOR1 - 1, 16)) And _
               (WS4.Cells(FOR1, 10) = WS4.Cells(FOR1 - 1, 10)) Then
                WS4.Cells(FOR1, 16).ClearContents

            End If
        Next FOR1

        For FOR1 = 2 To WS4_EC_LR

            If .CountIfs(WS4.Range("P:P"), WS4.Cells(FOR1, 16)) >= 2 Then
                WS4.Cells(FOR1, 27).Value = 1

            End If
        Next FOR1

        '在庫用の表作成
        For FOR1 = 2 To WS4_EC_LR
            WS4.Cells(FOR1, 12).Value = Trim(WS4.Cells(FOR1, 12))
        Next FOR1

        WS10.Range("A:B").Delete

        WS4.Range(WS4.Cells(1, 12), WS4.Cells(WS4_EC_LR, 12)).Copy
        WS10.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

        WS10.Range("A1:A" & WS4_EC_LR).RemoveDuplicates Columns:=1, Header:=xlYes

        Dim WS10_AC_LR As Long
        WS10_AC_LR = WS10.Cells(WS10.Rows.Count, 1).End(xlUp).Row

        WS10.Sort.SortFields.Clear
        WS10.Sort.SortFields.Add2 Key:=WS10.Range("A2:A" & WS10_AC_LR), _
                                  SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With WS10.Sort
            .SetRange WS10.Range("A1:A" & WS10_AC_LR)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        WS10.Cells(1, 2).Formula = "数量"
        WS10.Cells(2, 2).Formula = "=SUMIFS('アラジン取込用(売上)'!$R:$R,'アラジン取込用(売上)'!$L:$L,$A2)"
        WS10.Cells(2, 2).Copy
        WS10.Range(WS10.Cells(3, 2), WS10.Cells(WS10_AC_LR, 2)).PasteSpecial

        WS10.Columns("A:B").AutoFit

        '在庫用の表作成終了

        '運送便チェック用作成

        If WS11.AutoFilterMode = True Then
            WS11.AutoFilterMode = False
        End If

        WS11.Rows("2:3000").Delete

        WS4_EC_LR = WS4.Cells(WS4.Rows.Count, 5).End(xlUp).Row
        For FOR1 = 2 To WS4_EC_LR

            WS11.Cells(FOR1, 2).Value = WS4.Cells(FOR1, 8)
            If .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 5) = 0 Then
                WS11.Cells(FOR1, 7).Value = WS4.Cells(FOR1, 29)
            Else
                WS11.Cells(FOR1, 7).Value = WS4.Cells(FOR1, 29) & " 代引き"
            End If

            'WS11.Cells(FOR1, 3).Value = .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 25)
            'If WS11.Cells(FOR1, 3) <> WS11.Cells(FOR1 - 1, 3) Then
            WS11.Cells(FOR1, 3).Value = .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 27) _
                                      & .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 28) _
                                      & .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Range("AN:AN"), 0), 29)
            'End If

            WS11.Cells(FOR1, 10).Value = WS4.Cells(FOR1, 27)

            WS11.Cells(FOR1, 4).Value = WS4.Cells(FOR1, 12)
            WS11.Cells(FOR1, 5).Value = WS4.Cells(FOR1, 18)
            If .CountIf(WS11.Cells(FOR1, 3), "北海道*") >= 1 Then
                WS11.Range(WS11.Cells(FOR1, 3), WS11.Cells(FOR1, 3)).Interior.ColorIndex = 3
            ElseIf .SumIf(WS6.Range("B:B"), WS11.Cells(FOR1, 4), WS6.Range("F:F")) >= 1 Then
                WS11.Range(WS11.Cells(FOR1, 4), WS11.Cells(FOR1, 4)).Interior.ColorIndex = 3
            End If

            WS11.Cells(FOR1, 1).Value = Mid(.Index(WS9.Cells, .Match(WS4.Cells(FOR1, 5), WS9.Range("C:C"), 0), 5), 1, 1) _
                                      & " " & WS4.Cells(FOR1, 10) & " " & WS4.Cells(FOR1, 26)

        Next FOR1

        For FOR1 = 2 To WS4_EC_LR

            If .CountIf(WS11.Range(WS11.Cells(2, 2), WS11.Cells(FOR1, 2)), WS11.Cells(FOR1, 2)) >= 2 Then

                WS11.Cells(FOR1, 3).ClearContents
            End If

            If .CountIf(WS11.Range(WS11.Cells(2, 1), WS11.Cells(FOR1, 1)), WS11.Cells(FOR1, 1)) = 1 Then

                WS11.Cells(FOR1, 11).Value = WS11.Cells(FOR1, 7)
            End If

        Next FOR1

        For FOR1 = 2 To WS4_EC_LR

            If WS11.Cells(FOR1, 2) <> WS11.Cells(FOR1 + 1, 2) Then
                WS11.Range(WS11.Cells(FOR1, 1), WS11.Cells(FOR1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If

            If .CountIf(WS11.Cells(FOR1, 4), "Z-0*") = 1 Then
                WS11.Range(WS11.Cells(FOR1, 1), WS11.Cells(FOR1, 5)).Interior.ColorIndex = 1

            End If

            If .CountIf(WS11.Range("C:C"), WS11.Cells(FOR1, 3)) >= 2 Then
                WS11.Range(WS11.Cells(FOR1, 1), WS11.Cells(FOR1, 5)).Interior.ColorIndex = 6
            End If

            'If .CountIf(WS11.Range("B:B"), WS11.Cells(FOR1, 2)) >= 2 And _
             '.CountIf(WS11.Range("A:A"), WS11.Cells(FOR1, 1)) = 1 Then
            'WS11.Range(WS11.Cells(FOR1, 1), WS11.Cells(FOR1, 5)).Interior.ColorIndex = 6
            'End If

            If WS11.Cells(FOR1, 10) = 1 Then
                WS11.Range(WS11.Cells(FOR1, 1), WS11.Cells(FOR1, 7)).Interior.ColorIndex = 6
                WS11.Cells(FOR1, 7).Value = WS11.Cells(FOR1, 7) & " 注番違い同梱有り"
            End If

        Next FOR1

        For FOR1 = 2 To WS4_EC_LR
            If .CountIfs(WS4.Range(WS4.Cells(2, 16), WS4.Cells(FOR1, 16)), WS4.Cells(FOR1, 16)) >= 2 Then
                WS4.Cells(FOR1, 27).ClearContents
            End If
        Next FOR1

        With WS11.Columns("A:J").Font
            .Name = "メイリオ"
            .FontStyle = "レギュラー"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With

        WS11.Cells(1, 9).Value = "総件数 " & .Large(WS4.Range("P:P"), 1) - WS4.Cells(2, 16) + 1
        WS11.Cells(2, 9).Value = "内同梱数 " & .Sum(WS4.Range("AA:AA"))
        WS11.Cells(4, 9).Value = "ヤマト" & .CountIf(WS11.Range("K:K"), "ヤマト*") & "件"
        WS11.Cells(5, 9).Value = "クラフト" & .CountIf(WS11.Range("K:K"), "クラフト*") & "件"
        WS11.Cells(6, 9).Value = "ネコポス" & .CountIf(WS11.Range("K:K"), "ネコポス*") & "件"
        WS11.Cells(7, 9).Value = "佐川" & .CountIf(WS11.Range("K:K"), "佐川*") & "件"
        WS11.Range("J:K").Delete
        WS11.Columns("A:AA").AutoFit

        '運送便チェック表作成終了
        MsgBox "今回の件数は" & .Large(WS4.Range("P:P"), 1) - WS4.Cells(2, 16) + 1 & "件です" & vbCrLf & _
               "(内同梱" & .Sum(WS4.Range("AA:AA")) & "件)"
        'MsgBox "在庫用シートを印刷してください"
        WS4.Range("AA:AA").ClearContents

        WS4.Activate

        'WS4.Range("AA:AC").Delete
        'WS10.Copy
        'WS4.Copy

    End With
End Sub