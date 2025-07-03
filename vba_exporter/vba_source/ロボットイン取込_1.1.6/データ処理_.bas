Option Explicit

Sub データ処理()

    Dim WS1 As Worksheet
    Set WS1 = ThisWorkbook.Worksheets("データ貼付用")

    Dim WS2 As Worksheet
    Set WS2 = ThisWorkbook.Worksheets("データ抽出用")

    'Dim WS3 As Worksheet
    'Set WS3 = ThisWorkbook.Worksheets("アラジン取込用(受注)")

    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("アラジン取込用(売上)")

    Dim WS5 As Worksheet
    Set WS5 = ThisWorkbook.Worksheets("納品先マスタ")

    Dim WS6 As Worksheet
    Set WS6 = ThisWorkbook.Worksheets("商品変換マスタ")

    Dim WS9 As Worksheet
    Set WS9 = ThisWorkbook.Worksheets("モールマスタ")

    With WorksheetFunction

        If WS1.AutoFilterMode = True Then
            WS1.AutoFilterMode = False
        End If

        If WS2.AutoFilterMode = True Then
            WS2.AutoFilterMode = False
        End If

        'If WS3.AutoFilterMode = True Then
        'WS3.AutoFilterMode = False
        'End If

        If WS4.AutoFilterMode = True Then
            WS4.AutoFilterMode = False
        End If

        If WS5.AutoFilterMode = True Then
            WS5.AutoFilterMode = False
        End If

        WS2.Range("3:2000").ClearContents
        'WS3.Range("2:3000").Delete
        WS4.Range("2:3000").Delete

        Dim WS1_AC_LR As Long
        WS1_AC_LR = WS1.Cells(WS1.Rows.Count, 1).End(xlUp).Row

        Dim WS2_LC_1R As Long
        WS2_LC_1R = WS2.Cells(1, WS2.Columns.Count).End(xlToLeft).Column

        WS2.Range(WS2.Cells(2, 1), WS2.Cells(2, WS2_LC_1R)).Copy
        WS2.Range(WS2.Cells(2, 1), WS2.Cells(WS1_AC_LR, WS2_LC_1R)).PasteSpecial

        Dim WS2_AC_LR As Long
        WS2_AC_LR = WS2.Cells(WS2.Rows.Count, 1).End(xlUp).Row

        If WS1_AC_LR <= 2 Then
            Exit Sub
        End If

        '売上データ作成ここから

        Dim FOR1 As Long

        Dim WS4_AC_LR As Long

        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row

        '20221125 アマゾン AUPAY チェック用プログラム
        For FOR1 = 2 To WS2_AC_LR
            If ((WS1.Cells(FOR1, 22) + WS1.Cells(FOR1, 24)) <> 0) _
                And (WS1.Cells(FOR1, 3) = 2712 Or WS1.Cells(FOR1, 3) = 345715) Then
                MsgBox "アマゾンかauPAYで意図していない手数料もしくは割引が発生のため" _
                & vbCrLf & "コハラまで連絡してください"
            End If
        Next FOR1
        '20221125 アマゾン AUPAY チェック用プログラム終了

        For FOR1 = 2 To WS2_AC_LR

            'If (WS2.Cells(FOR1, 2) = 300) Or (WS2.Cells(FOR1, 2) = 200 And WS2.Cells(FOR1, 20) <> 0) Then

            If WS2.Cells(FOR1, 11) >= 1 Then    '注文入力

                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Value = WS2.Cells(FOR1, 8).Value    '商品CD
                WS4.Cells(WS4_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 10).Value    '売上数量
                WS4.Cells(WS4_AC_LR + 1, 19).Value = WS2.Cells(FOR1, 9).Value    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'01"    '取引区分
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            If WS2.Cells(FOR1, 3) > 0 Then    'クーポン処理

                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Value = "Z-0011-9"    '商品CD
                WS4.Cells(WS4_AC_LR + 1, 14).Value = "ｸｰﾎﾟﾝ値引"
                WS4.Cells(WS4_AC_LR + 1, 18).Value = "-1"   '売上数量
                WS4.Cells(WS4_AC_LR + 1, 19).Value = WS2.Cells(FOR1, 3).Value    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'03"    '取引区分 値引き
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            If WS2.Cells(FOR1, 12) > 0 Then    '決済手数料処理

                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Value = "Z-0001-1"    '商品CD
                WS4.Cells(WS4_AC_LR + 1, 14).Value = "決済手数料(ﾈｯﾄ)"
                WS4.Cells(WS4_AC_LR + 1, 18).Value = "1"   '売上数量
                WS4.Cells(WS4_AC_LR + 1, 19).Value = WS2.Cells(FOR1, 12).Value    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'01"    '取引区分
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            If WS2.Cells(FOR1, 4) > 0 Then    '送料処理

                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Value = "Z-0002-1"    '商品CD
                WS4.Cells(WS4_AC_LR + 1, 14).Value = "送料(ﾈｯﾄ)"
                WS4.Cells(WS4_AC_LR + 1, 18).Value = "1"   '売上数量
                WS4.Cells(WS4_AC_LR + 1, 19).Value = WS2.Cells(FOR1, 4).Value    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'01"    '取引区分
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            If WS2.Cells(FOR1, 5) > 0 Then    '代引き手数料処理
                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Value = "Z-0010-1"    '商品CD
                WS4.Cells(WS4_AC_LR + 1, 14).Value = "代引き手数料(ﾈｯﾄ)"
                WS4.Cells(WS4_AC_LR + 1, 18).Value = "1"   '売上数量
                WS4.Cells(WS4_AC_LR + 1, 19).Value = WS2.Cells(FOR1, 5).Value    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'01"    '取引区分
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            If WS2.Cells(FOR1, 37) <> "" Then  'プレゼント

                WS4.Cells(WS4_AC_LR + 1, 1).Value = WS2.Cells(FOR1, 19) * 1    '売上日付
                WS4.Cells(WS4_AC_LR + 1, 2).Value = WS2.Cells(FOR1, 19) * 1    '出荷日付
                WS4.Cells(WS4_AC_LR + 1, 10).Value = "'" & WS2.Cells(FOR1, 13)    '相手先No
                WS4.Cells(WS4_AC_LR + 1, 12).Interior.ColorIndex = 3    '入力用赤塗りつぶし
                'On Error Resume Next
                WS4.Cells(WS4_AC_LR + 1, 12).Value = WS2.Cells(FOR1, 37)
                'On Error GoTo 0

                WS4.Cells(WS4_AC_LR + 1, 14).Interior.ColorIndex = 3

                WS4.Cells(WS4_AC_LR + 1, 18).Interior.ColorIndex = 3
                WS4.Cells(WS4_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 38)
                WS4.Cells(WS4_AC_LR + 1, 19).Value = 0    '売上金額
                WS4.Cells(WS4_AC_LR + 1, 5).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 3)    '得意先CD
                WS4.Cells(WS4_AC_LR + 1, 7).Value = .Index(WS9.Cells, .Match(WS2.Cells(FOR1, 33), WS9.Range("A:A"), 0), 4)    '納品先CD
                WS4.Cells(WS4_AC_LR + 1, 9).Value = "'00000001"    '倉庫
                WS4.Cells(WS4_AC_LR + 1, 11).Value = "'01"    '取引区分
                WS4.Cells(WS4_AC_LR + 1, 24).Value = WS2.Cells(FOR1, 34) * 1    '配送希望日

                If WS2.Cells(FOR1, 14) <> 1 Then
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                    WS4.Cells(WS4_AC_LR + 1, 26).Value = WS2.Cells(FOR1, 15)    '注文者
                Else
                    WS4.Cells(WS4_AC_LR + 1, 25).Value = WS2.Cells(FOR1, 17)    '送付先
                End If

                WS4_AC_LR = WS4_AC_LR + 1

            End If

            'End If
            '-----------------------------------------20221122ここまで
        Next FOR1

        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row

        '同一人物別の楽天ショップの場合は同梱処理しない様に変更 20230703
        For FOR1 = 2 To WS4_AC_LR

            WS4.Cells(FOR1, 28).Value = _
            .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 17) & _
            .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 27) & _
            .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 28) & _
            .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 29) & _
            .Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 33)

            WS4.Cells(FOR1, 30).Value = _
            .VLookup(.Index(WS2.Cells, .Match(WS4.Cells(FOR1, 10) * 1, WS2.Columns("AP"), 0), 33), WS9.Range("A:B"), 2, False)

        Next FOR1

        'セット品番用変換 240826

        Dim result As Variant
        Dim lastRow As Long
        Dim str1 As Variant
        Dim str2 As Variant
        Dim str3 As Variant

        ' 最終行を取得
        lastRow = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row

        FOR1 = 2
        str1 = "OSHI-NB2914"
        str2 = "OSHI-NB2520"
        str3 = "OSHI-NB3320"

        Do While FOR1 <= lastRow

            If Mid(WS4.Cells(FOR1, "L"), 1, 11) = str1 Or Mid(WS4.Cells(FOR1, "L"), 1, 11) = str2 Or Mid(WS4.Cells(FOR1, "L"), 1, 11) = str3 Then
                result = Application.Match(WS4.Cells(FOR1, "L"), WS6.Range("A:A"), 0)
                If IsError(result) Then

                    MsgBox WS4.Cells(FOR1, "L") & " 商品コードを登録してください。"
                    Exit Sub
                End If
            End If

            FOR1 = FOR1 + 1
        Loop

        FOR1 = 2

        Do While FOR1 <= lastRow
            ' VLookupの結果を変数に格納し、エラーが発生しないか確認
            result = Application.VLookup(WS4.Cells(FOR1, "L"), WS6.Range("A:M"), 8, False)

            'Debug.Print .VLookup(WS4.Cells(FOR1, "L"), WS6.Range("A:M"), 7, False) * 1
            ' 結果がエラーでない場合に処理を実行
            If Not IsError(result) Then

                ' VLookupがエラーを返さなかった場合の処理

                If Mid(WS4.Cells(FOR1, "L"), 1, 11) = str1 Or Mid(WS4.Cells(FOR1, "L"), 1, 11) = str2 Or Mid(WS4.Cells(FOR1, "L"), 1, 11) = str3 Then

                    ' FOR1行をコピーして、次の行に挿入
                    WS4.Rows(FOR1).Copy
                    WS4.Rows(FOR1 + 1).Insert Shift:=xlDown
                    WS4.Rows(FOR1 + 1).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                    WS4.Cells(FOR1, 19).Value = .VLookup(WS4.Cells(FOR1, "L"), WS6.Range("A:M"), 11, False)
                    WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, "L"), WS6.Range("A:M"), 8, False)

                    WS4.Cells(FOR1 + 1, 19).Value = .VLookup(WS4.Cells(FOR1 + 1, "L"), WS6.Range("A:M"), 12, False)
                    WS4.Cells(FOR1 + 1, 12).Value = .VLookup(WS4.Cells(FOR1 + 1, "L"), WS6.Range("A:M"), 9, False)

                    ' 最終行を更新して、新しい行を考慮
                    lastRow = lastRow + 1

                    ' FOR1を次の行に進める（挿入した行をスキップする）
                    FOR1 = FOR1 + 1

                End If
            End If

            ' 次の行に進める
            FOR1 = FOR1 + 1
        Loop

        '-----------------------------------------240826ここまで

        '受注データ作成ここから

        'Dim WS3_AC_LR As Long
        '
        'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
        '
        'For FOR1 = 2 To WS2_AC_LR
        '
        ''If (WS2.Cells(FOR1, 2) = 100) Or (WS2.Cells(FOR1, 2) = 200 And WS2.Cells(FOR1, 20) = 0) Then
        '
        'If WS2.Cells(FOR1, 11) >= 1 Then '注文入力
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Value = WS2.Cells(FOR1, 8).Value '商品CD
        'WS3.Cells(WS3_AC_LR + 1, 17).Value = WS2.Cells(FOR1, 10).Value '売上数量
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 9).Value '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        'If WS2.Cells(FOR1, 21) > 0 Then 'クーポン処理
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0011-9" '商品CD
        'WS3.Cells(WS3_AC_LR + 1, 13).Value = "ｸｰﾎﾟﾝ値引"
        'WS3.Cells(WS3_AC_LR + 1, 17).Value = "-1"   '売上数量
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 21).Value '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分 値引き
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        'If WS2.Cells(FOR1, 12) > 0 Then '決済手数料処理
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0001-1" '商品CD
        'WS3.Cells(WS3_AC_LR + 1, 13).Value = "決済手数料(ﾈｯﾄ)"
        'WS3.Cells(WS3_AC_LR + 1, 17).Value = "1"   '売上数量
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 12).Value '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        'If WS2.Cells(FOR1, 4) > 0 Then '送料処理
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0002-1" '商品CD
        'WS3.Cells(WS3_AC_LR + 1, 13).Value = "送料(ﾈｯﾄ)"
        'WS3.Cells(WS3_AC_LR + 1, 17).Value = "1"   '売上数量
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 4).Value '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        'If WS2.Cells(FOR1, 5) > 0 Then '代引き手数料処理
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Value = "Z-0010-1" '商品CD
        'WS3.Cells(WS3_AC_LR + 1, 13).Value = "代引き手数料(ﾈｯﾄ)"
        'WS3.Cells(WS3_AC_LR + 1, 17).Value = "1"   '売上数量
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = WS2.Cells(FOR1, 5).Value '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        'If .CountIf(WS2.Cells(FOR1, 23), "*プレゼント*") > 0 Then 'プレゼント
        '
        'WS3.Cells(WS3_AC_LR + 1, 1).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用受注日付
        'WS3.Cells(WS3_AC_LR + 1, 2).Value = Mid(WS2.Cells(FOR1, 1), 8, 8) '受注データ用指定納期
        'WS3.Cells(WS3_AC_LR + 1, 9).Value = "'" & Mid(WS2.Cells(FOR1, 1), 17, 10) '相手先No
        'WS3.Cells(WS3_AC_LR + 1, 11).Interior.ColorIndex = 3
        'WS3.Cells(WS3_AC_LR + 1, 13).Interior.ColorIndex = 3
        'WS3.Cells(WS3_AC_LR + 1, 17).Interior.ColorIndex = 3
        'WS3.Cells(WS3_AC_LR + 1, 18).Value = 0 '売上金額
        'WS3.Cells(WS3_AC_LR + 1, 4).Value = "79999003" '得意先CD
        'WS3.Cells(WS3_AC_LR + 1, 6).Value = "80013313" '納品先CD
        'WS3.Cells(WS3_AC_LR + 1, 8).Value = "'00000001" '倉庫
        'WS3.Cells(WS3_AC_LR + 1, 10).Value = "'01" '取引区分
        '
        'If WS2.Cells(FOR1, 14) <> 1 Then
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) _
        '& " (注) " & WS2.Cells(FOR1, 15) & " " & WS2.Cells(FOR1, 16) '送付先＆注文者
        'Else
        'WS3.Cells(WS3_AC_LR + 1, 23).Value = WS2.Cells(FOR1, 17) & " " & WS2.Cells(FOR1, 18) '送付先
        'End If
        '
        'WS3_AC_LR = WS3_AC_LR + 1
        '
        'End If
        '
        ''End If
        '
        'Next FOR1

        'ここから商品コード変換処理

        'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
        '
        'For FOR1 = 2 To WS3_AC_LR
        '
        'If .CountIf(WS6.Range("A:A"), WS3.Cells(FOR1, 11)) >= 1 Then
        'If .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:D"), 4, False) = 1 Then
        'WS3.Range(WS3.Cells(FOR1, 1), WS3.Cells(FOR1, 26)).Interior.ColorIndex = 3
        'End If
        '
        '
        'If .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:C"), 3, False) = "" Then
        'WS3.Cells(FOR1, 11).Value = .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:D"), 2, False)
        'Else
        'WS3.Cells(FOR1, 17).Value = WS3.Cells(FOR1, 17) * .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:C"), 3, False)
        'WS3.Cells(FOR1, 18).Value = WS3.Cells(FOR1, 18) / .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:C"), 3, False)
        'WS3.Cells(FOR1, 11).Value = .VLookup(WS3.Cells(FOR1, 11), WS6.Range("A:D"), 2, False)
        'End If
        'End If
        '
        'Next FOR1

        '例外処理ここから
        '        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '
        '        For FOR1 = 2 To WS4_AC_LR
        '
        '            'WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '
        '            '要注意商品処理・・・手作業など
        '            If .CountIf(WS6.Range("A:A"), WS4.Cells(FOR1, 12)) >= 1 Then
        '                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 4, False) = 1 Then
        '                    WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 26)).Interior.ColorIndex = 3
        '                End If
        '
        '                'セット商品単品変更処理
        '                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False) = "" Then    'セット商品ではなく単純な登録ミスの場合
        '                    WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                Else    'セット商品の場合
        '                    WS4.Cells(FOR1, 20).Value = WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
        '                    WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18) * .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
        '                    If (WS4.Cells(FOR1, 19) Mod .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)) = 0 Then    'セット商品で割り切れる場合
        '                        WS4.Cells(FOR1, 19).Value = WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
        '                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                        WS4.Cells(FOR1, 20).ClearContents
        '                    Else
        '
        '                        Dim SYOUHINKEI As Long    'セット商品で割り切れない場合
        '                        SYOUHINKEI = WS4.Cells(FOR1, 20).Value
        '                        WS4.Cells(FOR1, 19).Value = .Ceiling(WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False), 1)
        '                        WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18).Value - 1
        '                        WS4.Rows(FOR1 + 1).Insert
        '                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                        WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 17)).Copy
        '                        WS4.Range(WS4.Cells(FOR1 + 1, 1), WS4.Cells(FOR1 + 1, 17)).PasteSpecial
        '                        WS4.Range(WS4.Cells(FOR1, 20), WS4.Cells(FOR1, 26)).Copy
        '                        WS4.Range(WS4.Cells(FOR1 + 1, 20), WS4.Cells(FOR1 + 1, 26)).PasteSpecial
        '                        WS4.Cells(FOR1 + 1, 18).Value = 1
        '                        WS4.Cells(FOR1 + 1, 19).Value = SYOUHINKEI - WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
        '                        WS4.Cells(FOR1 + 1, 20).ClearContents
        '                        WS4.Cells(FOR1, 20).ClearContents
        '                        FOR1 = FOR1 + 1
        '                        WS4_AC_LR = WS4_AC_LR + 1
        '                    End If
        '                End If
        '            End If
        '            'WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '        Next FOR1
        '
        '        '20221227最終行が抜けるのでとりあえず解決策見当たらず2回Forを回す
        '        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '
        '        For FOR1 = 2 To WS4_AC_LR
        '
        '            'WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '
        '            '要注意商品処理・・・手作業など
        '            If .CountIf(WS6.Range("A:A"), WS4.Cells(FOR1, 12)) >= 1 Then
        '                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 4, False) = 1 Then
        '                    WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 26)).Interior.ColorIndex = 3
        '                End If
        '
        '                'セット商品単品変更処理
        '                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False) = "" Then    'セット商品ではなく単純な登録ミスの場合
        '                    WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                Else    'セット商品の場合
        '                    WS4.Cells(FOR1, 20).Value = WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
        '                    WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18) * .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
        '                    If (WS4.Cells(FOR1, 19) Mod .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)) = 0 Then    'セット商品で割り切れる場合
        '                        WS4.Cells(FOR1, 19).Value = WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
        '                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                        WS4.Cells(FOR1, 20).ClearContents
        '                    Else
        '
        '                        'Dim SYOUHINKEI As Long 'セット商品で割り切れない場合
        '                        SYOUHINKEI = WS4.Cells(FOR1, 20).Value
        '                        WS4.Cells(FOR1, 19).Value = .Ceiling(WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False), 1)
        '                        WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18).Value - 1
        '                        WS4.Rows(FOR1 + 1).Insert
        '                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
        '                        WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 17)).Copy
        '                        WS4.Range(WS4.Cells(FOR1 + 1, 1), WS4.Cells(FOR1 + 1, 17)).PasteSpecial
        '                        WS4.Range(WS4.Cells(FOR1, 20), WS4.Cells(FOR1, 26)).Copy
        '                        WS4.Range(WS4.Cells(FOR1 + 1, 20), WS4.Cells(FOR1 + 1, 26)).PasteSpecial
        '                        WS4.Cells(FOR1 + 1, 18).Value = 1
        '                        WS4.Cells(FOR1 + 1, 19).Value = SYOUHINKEI - WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
        '                        WS4.Cells(FOR1 + 1, 20).ClearContents
        '                        WS4.Cells(FOR1, 20).ClearContents
        '                        FOR1 = FOR1 + 1
        '                        WS4_AC_LR = WS4_AC_LR + 1
        '                    End If
        '                End If
        '            End If
        '            'WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        '        Next FOR1

        'Dim WS4_AC_LR As Long
        'Dim FOR1 As Long

        '240827上記のコードをリファクタリング
        Dim SYOUHINKEI As Long

        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row

        For FOR1 = WS4_AC_LR To 2 Step -1

            If .CountIf(WS6.Range("A:A"), WS4.Cells(FOR1, 12)) >= 1 Then
                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 4, False) = 1 Then
                    WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 26)).Interior.ColorIndex = 3
                End If

                If .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False) = "" Then
                    WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
                Else
                    WS4.Cells(FOR1, 20).Value = WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
                    WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18) * .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
                    If (WS4.Cells(FOR1, 19) Mod .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)) = 0 Then
                        WS4.Cells(FOR1, 19).Value = WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False)
                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
                        WS4.Cells(FOR1, 20).ClearContents
                    Else
                        SYOUHINKEI = WS4.Cells(FOR1, 20).Value
                        WS4.Cells(FOR1, 19).Value = .Ceiling(WS4.Cells(FOR1, 19) / .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:C"), 3, False), 1)
                        WS4.Cells(FOR1, 18).Value = WS4.Cells(FOR1, 18).Value - 1
                        WS4.Rows(FOR1 + 1).Insert
                        WS4.Cells(FOR1, 12).Value = .VLookup(WS4.Cells(FOR1, 12), WS6.Range("A:D"), 2, False)
                        WS4.Range(WS4.Cells(FOR1, 1), WS4.Cells(FOR1, 17)).Copy
                        WS4.Range(WS4.Cells(FOR1 + 1, 1), WS4.Cells(FOR1 + 1, 17)).PasteSpecial
                        WS4.Range(WS4.Cells(FOR1, 20), WS4.Cells(FOR1, 26)).Copy
                        WS4.Range(WS4.Cells(FOR1 + 1, 20), WS4.Cells(FOR1 + 1, 26)).PasteSpecial
                        WS4.Cells(FOR1 + 1, 18).Value = 1
                        WS4.Cells(FOR1 + 1, 19).Value = SYOUHINKEI - WS4.Cells(FOR1, 18) * WS4.Cells(FOR1, 19)
                        WS4.Cells(FOR1 + 1, 20).ClearContents
                        WS4.Cells(FOR1, 20).ClearContents
                    End If
                End If
            End If
        Next FOR1
        '-----------------------------------------240827ここまで

        '20221125 プレゼント品番 チェック用プログラム
        WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
        For FOR1 = 2 To WS4_AC_LR
            If (WS4.Cells(FOR1, 19) = 0) And (Mid(WS4.Cells(FOR1, 12), 1, 4) <> "MOC-") Then
                If Mid(WS4.Cells(FOR1, 12), 1, 6) <> "WEB-BS" Then
                    If Mid(WS4.Cells(FOR1, 12), 1, 10) <> "WEB-HINOKI" Then
                        If Mid(WS4.Cells(FOR1, 12), 1, 5) <> "要チェック" Then
                            MsgBox "プレゼント品番が不正です。EC事業部に確認"
                        End If
                    End If
                End If
            End If

            'If ((Mid(WS4.Cells(FOR1, 12), 1, 8) <> "MOC-RPTS") Or (Mid(WS4.Cells(FOR1, 12), 1, 6) <> "WEB-BS") _
            'Or (Mid(WS4.Cells(FOR1, 12), 1, 10) <> "WEB-HINOKI")) And (WS4.Cells(FOR1, 19) = 0) Then
            'MsgBox "プレゼント品番が不正です。EC事業部に確認"
            'End If
        Next FOR1
        '20221125 プレゼント品番 チェック用プログラム終了

        For FOR1 = 2 To WS4_AC_LR
            If WS4.Cells(FOR1, 28) = "" Then
                WS4.Cells(FOR1, 28).Value = WS4.Cells(FOR1 - 1, 28)
            End If
        Next FOR1

        WS4.Activate

    End With

End Sub