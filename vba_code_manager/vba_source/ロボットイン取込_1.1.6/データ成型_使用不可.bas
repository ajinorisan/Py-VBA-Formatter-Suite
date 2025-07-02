Sub データ成型()
    Application.ScreenUpdating = False

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("データ貼付用")

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("データ抽出用")

    Dim WS3 As Worksheet
    Set WS3 = ThisWorkbook.Worksheets("アラジン取込用(受注)")

    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("アラジン取込用(売上)")

    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("納品先マスタ")

    With WorksheetFunction

        If WS3.AutoFilterMode = True Then
            WS3.AutoFilterMode = False
        End If
        If WS4.AutoFilterMode = True Then
            WS4.AutoFilterMode = False
        End If

        WS2.Range("2:2000").ClearContents
        WS3.Range("2:3000").Delete
        WS4.Range("2:3000").Delete

        Dim WS1_AC_LR As Long
        WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

        Dim WS2_LC_1R As Long
        WS2_LC_1R = WS2.Cells(1, WS2.Columns.Count).End(xlToLeft).Column

        WS2.Range(WS2.Cells(1, 1), WS2.Cells(1, WS2_LC_1R)).Copy
        WS2.Range(WS2.Cells(2, 1), WS2.Cells(WS1_AC_LR, WS2_LC_1R)).PasteSpecial

        Dim WS2_AC_LR As Long
        WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row

        'Dim WS2_COUNT As Long
        'WS2_COUNT = WorksheetFunction.Sum(WS2.Range("K:K")) + 1

        Dim FOR1 As Long

        Dim WS3_AC_LR As Long

        For FOR1 = 2 To WS2_AC_LR

            WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
            'For FOR1 = WS3_AC_LR + 1 To WS2_COUNT
            If WS2.Cells(FOR1, 11) = 1 Then 'クーポン送料代引きが無い行の処理

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = WS2.Cells(FOR1, 8).Value
                WS3.Cells(WS3_AC_LR + 1, 17).Value = WS2.Cells(FOR1, 10).Value
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 9).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

            End If

            WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

            If WS2.Cells(FOR1, 11) >= 2 Then

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 8)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = WS2.Cells(FOR1, 8).Value
                WS3.Cells(WS3_AC_LR + 1, 17).Value = WS2.Cells(FOR1, 10).Value
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 9).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

                WS3_AC_LR = WS3_AC_LR + 1

            End If

            If (WS2.Cells(FOR1, 3) > 0) Then  'クーポン値引き処理

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 8)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0011-9"
                WS3.Cells(WS3_AC_LR + 1, 13).Value = "ｸｰﾎﾟﾝ値引"
                WS3.Cells(WS3_AC_LR + 1, 17).Value = "-1"
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 3).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'02"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

                WS3_AC_LR = WS3_AC_LR + 1

            End If

            If (WS2.Cells(FOR1, 4) > 0) Then  '送料処理

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 8)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0002-1"
                WS3.Cells(WS3_AC_LR + 1, 13).Value = "送料(ﾈｯﾄ)"
                WS3.Cells(WS3_AC_LR + 1, 17).Value = 1
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 4).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

                WS3_AC_LR = WS3_AC_LR + 1

            End If

            If (WS2.Cells(FOR1, 5) > 0) Then '代引き手数料処理

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 8)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0010-1"
                WS3.Cells(WS3_AC_LR + 1, 13).Value = "代引き手数料(ﾈｯﾄ)"
                WS3.Cells(WS3_AC_LR + 1, 17).Value = 1
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 5).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

                WS3_AC_LR = WS3_AC_LR + 1

            End If

            If (WS2.Cells(FOR1, 12) > 0) Then '決済手数料処理

                WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8)
                WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 8)
                WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0001-1"
                WS3.Cells(WS3_AC_LR + 1, 13).Value = "決済手数料(ﾈｯﾄ)"
                WS3.Cells(WS3_AC_LR + 1, 17).Value = 1
                WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 12).Value
                WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003"
                WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313"
                WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001"
                WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01"

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
                    & "(注)" & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16)
                Else
                    WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18)
                End If

                WS3_AC_LR = WS3_AC_LR + 1

            End If

        Next FOR1

        WS3.Range("X:AA").Delete

        Application.ScreenUpdating = True

        WS3.Activate

        Application.ScreenUpdating = False

        受注取込振分:

        On Error GoTo 0

        Dim IB As String
        IB = Application.InputBox(Prompt:="受注入力する受注No?(半角数字)" & vbCrLf & _
        "複数ある場合は-(ハイフン)で区切ってください" _
        & vbCrLf & " 2020/05/26現在 4注番まで対応(増やす必要あるか検証中)", Type:=2)

        If IB = "" Then GoTo 売上取込振分
        If IB = "False" Then GoTo 売上取込振分

        WS3.Range("X1").Value = StrConv(IB, vbNarrow)

        If Len(WS3.Range("X1")) - Len(.Substitute(WS3.Range("X1"), "-", "")) = 0 Then
            WS3.Range("Y1").Value = WS3.Range("X1").Value
        End If

        If Len(WS3.Range("X1")) - Len(.Substitute(WS3.Range("X1"), "-", "")) = 1 Then
            WS3.Range("Y1").Value = _
            Mid(WS3.Range("X1"), 1, .Find("-", WS3.Range("X1"), 1) - 1)
            WS3.Range("Y2").Value = _
            Mid(WS3.Range("X1"), .Find("-", WS3.Range("X1"), 1) + 1, 8)
        End If

        If Len(WS3.Range("X1")) - Len(.Substitute(WS3.Range("X1"), "-", "")) = 2 Then
            WS3.Range("X2").Value = _
            Mid(WS3.Range("X1"), .Find("-", WS3.Range("X1")) + 1, 17)
            WS3.Range("X1").Value = Mid(WS3.Range("X1"), 1, .Find("-", WS3.Range("X1")) - 1)
            WS3.Range("Y1").Value = WS3.Range("X1").Value
            WS3.Range("Y2").Value = _
            Mid(WS3.Range("X2"), 1, .Find("-", WS3.Range("X2")) - 1)
            WS3.Range("Y3").Value = _
            Mid(WS3.Range("X2"), .Find("-", WS3.Range("X2")) + 1, 8)
        End If

        If Len(WS3.Range("X1")) - Len(.Substitute(WS3.Range("X1"), "-", "")) = 3 Then
            WS3.Range("X2").Value = _
            Mid(WS3.Range("X1"), .Find("-", WS3.Range("X1")) + 1, 26)
            WS3.Range("X3").Value = _
            Mid(WS3.Range("X2"), .Find("-", WS3.Range("X2")) + 1, 17)

            WS3.Range("X1").Value = Mid(WS3.Range("X1"), 1, .Find("-", WS3.Range("X1")) - 1)
            WS3.Range("X2").Value = Mid(WS3.Range("X2"), 1, .Find("-", WS3.Range("X2")) - 1)

            WS3.Range("Y1").Value = WS3.Range("X1").Value
            WS3.Range("Y2").Value = WS3.Range("X2").Value
            WS3.Range("Y3").Value = _
            Mid(WS3.Range("X3"), 1, .Find("-", WS3.Range("X3")) - 1)
            WS3.Range("Y4").Value = _
            Mid(WS3.Range("X3"), .Find("-", WS3.Range("X3")) + 1, 8)
        End If

        'If .CountIf(WS3.Range("X1"), "*.*") = 2 Then
        'WS3.Range("Y1").Value = _
        'Mid(WS3.Range("X1"), 1, .Find("-", WS3.Range("X1")) - 1)

        'End If

        WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

        For FOR1 = 1 To WS3_AC_LR
            On Error Resume Next
            If WS3.Cells(FOR1, 25) <> Int(WS3.Cells(FOR1, 25)) Then
                WS3.Range("X:AA").Delete
                GoTo 受注取込振分
            End If
        Next FOR1

        For FOR1 = 2 To WS3_AC_LR
            WS3.Cells(FOR1, 26).Value = WS3.Cells(FOR1, 9) * 1
        Next FOR1

        For FOR1 = 2 To WS3_AC_LR

            WS3.Cells(FOR1, 27).Value = _
            WorksheetFunction.CountIf(WS3.Range("Y:Y"), WS3.Cells(FOR1, 26))
        Next FOR1

        'WS3.Copy
        売上取込振分:

        WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

        For FOR1 = 2 To WS3_AC_LR
            'If WS3.Cells(FOR1, 27) <> "" Then
            WS3.Cells(FOR1, 27).Value = .Substitute(WS3.Cells(FOR1, 27), "0", "")
            'End If
        Next FOR1

        売上日付入力:
        Dim IB2 As String
        IB2 = Application.InputBox(Prompt:="売上日付入力(半角数字8桁/YYYYMMDD)" & vbCrLf & _
        "例)20200525", Type:=2)
        If IB2 = "" Then GoTo 売上日付入力
        If IB2 = "False" Then GoTo EXITSUB
        WS3.Range("Z1").Value = IB2
        If Len(WS3.Range("Z1")) <> 8 Then GoTo 売上日付入力

        'WS3.Range("W1").Value = IB2

        希望納期入力:
        Dim IB3 As String
        IB3 = Application.InputBox(Prompt:="希望納期入力(半角数字8桁/YYYYMMDD)" & vbCrLf & _
        "例)20200526", Type:=2)
        If IB3 = "" Then GoTo 希望納期入力
        If IB3 = "False" Then GoTo EXITSUB
        WS3.Range("Z2").Value = IB3
        If Len(WS3.Range("Z2")) <> 8 Then GoTo 希望納期入力

        納品先CD入力:
        Dim IB4 As String
        IB4 = Application.InputBox(Prompt:="納品先CD個人No?", Type:=2)
        If IB4 = "" Then GoTo 納品先CD入力
        If IB4 = "False" Then GoTo EXITSUB
        WS3.Range("Z3").Value = IB4
        If Len(WS3.Range("Z3")) >= 4 Then GoTo 納品先CD入力

        'If WS3.Cells(FOR1, 27) <> "" Then
        WS3.Range("AA1").Formula = "-"
        WS3.Range("A1:AA1").AutoFilter Field:=27, Criteria1:=""
        'End If
        WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

        WS3.Range(WS3.Cells(2, 4), WS3.Cells(WS3_AC_LR, 4)).Copy
        WS4.Range(WS4.Cells(2, 5), WS4.Cells(WS3_AC_LR, 5)) _
        .PasteSpecial Paste:=xlPasteValues

        WS3.Range(WS3.Cells(2, 8), WS3.Cells(WS3_AC_LR, 13)).Copy
        WS4.Range(WS4.Cells(2, 9), WS4.Cells(WS3_AC_LR, 14)) _
        .PasteSpecial Paste:=xlPasteValues

        WS3.Range(WS3.Cells(2, 17), WS3.Cells(WS3_AC_LR, 18)).Copy
        WS4.Range(WS4.Cells(2, 18), WS4.Cells(WS3_AC_LR, 19)) _
        .PasteSpecial Paste:=xlPasteValues

        'WS3.Range(WS3.Cells(2, 23), WS3.Cells(WS3_AC_LR, 23)).Copy
        'WS4.Range(WS4.Cells(2, 25), WS4.Cells(WS3_AC_LR, 25)) _
        '.PasteSpecial Paste:=xlPasteValues

        Dim WS4_EC_LR As Long
        WS4_EC_LR = WS4.Cells(WS4.Rows.Count, 5).End(xlUp).Row

        WS3.Range("Z1").Copy
        WS4.Range(WS4.Cells(2, 1), WS4.Cells(WS4_EC_LR, 2)) _
        .PasteSpecial Paste:=xlPasteValues

        WS3.Range("Z2").Copy
        WS4.Range(WS4.Cells(2, 24), WS4.Cells(WS4_EC_LR, 24)) _
        .PasteSpecial Paste:=xlPasteValues

        WS4.Cells(2, 7).Value = .Index(WS5.Cells, .Match(WS3.Range("Z3"), WS5.Range("E:E"), 0), 1)

        For FOR1 = 2 To WS4_EC_LR

            WS4.Cells(FOR1, 25).Value = _
            .VLookup(WS4.Cells(FOR1, 10), WS2.Range("M:R"), 5, False) _
            & " " & .VLookup(WS4.Cells(FOR1, 10), WS2.Range("M:R"), 6, False)

            If .VLookup(WS4.Cells(FOR1, 10), WS2.Range("M:R"), 2, False) <> 1 Then
                WS4.Cells(FOR1, 26).Value = _
                .VLookup(WS4.Cells(FOR1, 10), WS2.Range("M:R"), 3, False) _
                & " " & .VLookup(WS4.Cells(FOR1, 10), WS2.Range("M:R"), 4, False)
            End If
        Next FOR1

        Dim 納品個人番号 As Long
        納品個人番号 = WS3.Range("Z3")

        For FOR1 = 3 To WS4_EC_LR
            If WS4.Cells(FOR1, 10) <> WS4.Cells(FOR1 - 1, 10) Then
                WS4.Cells(FOR1, 7).Value = .Index(WS5.Cells, .Match(納品個人番号 + 1, WS5.Range("E:E"), 0), 1)
                納品個人番号 = 納品個人番号 + 1
            Else
                WS4.Cells(FOR1, 7).Value = WS4.Cells(FOR1 - 1, 7)
            End If
        Next FOR1

        If .Sum(WS3.Range("AA:AA")) >= 1 Then

            WS3.Range("2:" & WS3_AC_LR).Delete
            WS3.AutoFilterMode = False
            WS3.Range("X:AA").Delete
            'WS3.Copy

            'ActiveWorkbook.Worksheets("アラジン取込用(受注)").OLEObjects.Delete
            'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
            'For FOR1 = 2 To WS3_AC_LR
            'If ActiveWorkbook.Worksheets("アラジン取込用(受注)").Cells(FOR1, 10) = "02" Then
            'ActiveWorkbook.Worksheets("アラジン取込用(受注)").Cells(FOR1, 10).Formula = "'01"
            'End If
            'Next FOR1

            'WS4.Copy
        Else
            WS3.Range("X:AA").Delete
            'WS4.Copy
        End If

    End With

    EXITSUB:
    WS3.Range("X:AA").Delete
    'End With
    Application.ScreenUpdating = True
    Exit Sub

    Application.ScreenUpdating = True

End Sub