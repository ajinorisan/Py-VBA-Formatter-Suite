
Option Explicit

Private WSAO取込用 As Worksheet
Private WSmst As Worksheet
Private ws貼付 As Worksheet
Private WS計算 As Worksheet

Private ループ As Long

Private Sub マスタシート再計算()

    Set WSmst = ThisWorkbook.Worksheets("mst")

    For ループ = 1 To 最終行取得(WSmst, 1)
        WSmst.Cells(ループ, 1) = WSmst.Cells(ループ, 1).Formula
    Next ループ
End Sub

Private Sub 貼付シート整形()

    Set ws貼付 = ThisWorkbook.Worksheets("貼付")
    If InStr(ws貼付.Cells(1, 1).Value, "様") > 0 Then
    
        ws貼付.Rows("1:1").Delete
    End If
End Sub

Private Sub 計算シート作成()

    Set WSmst = ThisWorkbook.Worksheets("mst")
    Set ws貼付 = ThisWorkbook.Worksheets("貼付")
    Set WS計算 = ThisWorkbook.Worksheets("計算")

    Call 行削除(WS計算, 1, 最終行取得(WS計算, 1))

    WS計算.Cells(1, 1) = ws貼付.Cells(3, 3).Value

    With WorksheetFunction

        For ループ = 3 To 最終行取得(ws貼付, 3)

            ws貼付.Cells(ループ, 3) = ws貼付.Cells(ループ, 3).Formula

            If .CountIf(WS計算.Range("A:A"), ws貼付.Cells(ループ, 3)) <> 1 Then

                WS計算.Cells(最終行取得(WS計算, 1) + 1, 1) = ws貼付.Cells(ループ, 3).Value

            End If

        Next ループ

        Dim WSmst範囲 As Range
        Set WSmst範囲 = WSmst.Range("A:D")

        Dim WS貼付条件範囲 As Range
        Set WS貼付条件範囲 = ws貼付.Range("C:C")

        Dim WS貼付合計範囲 As Range
        Set WS貼付合計範囲 = ws貼付.Range("H:H")

        For ループ = 1 To 最終行取得(WS計算, 1)

            WS計算.Cells(ループ, 2) = .VLookup(WS計算.Cells(ループ, 1), WSmst範囲, 2, False)
            WS計算.Cells(ループ, 3) = .VLookup(WS計算.Cells(ループ, 1), WSmst範囲, 3, False)
            WS計算.Cells(ループ, 4) = .VLookup(WS計算.Cells(ループ, 1), WSmst範囲, 4, False)

            WS計算.Cells(ループ, 5) = .SumIfs(WS貼付合計範囲, WS貼付条件範囲, WS計算.Cells(ループ, 1))
            WS計算.Cells(ループ, 6) = -(.SumIfs(WS貼付合計範囲, WS貼付条件範囲, WS計算.Cells(ループ, 1)))
            WS計算.Cells(ループ, 7) = ((WS計算.Cells(ループ, 3) - WS計算.Cells(ループ, 4)) * WS計算.Cells(ループ, 5))

        Next ループ
    End With
End Sub

Private Sub 取込用シート作成()

    Set WSAO取込用 = ThisWorkbook.Worksheets("AO取込用")
    Set ws貼付 = ThisWorkbook.Worksheets("貼付")
    Set WS計算 = ThisWorkbook.Worksheets("計算")

    If 最終行取得(WSAO取込用, 12) <> 1 Then
        Call 行削除(WSAO取込用, 2, 最終行取得(WSAO取込用, 12))
    End If

    Dim 日付入力 As Long

    日付入力 = InputBox("20240831 の形で日付入力")
    WSAO取込用.Cells(2, 1).Value = 日付入力

    Dim WSAO取込用最終行 As Long
    WSAO取込用最終行 = 最終行取得(WSAO取込用, 12)    ' ループの前に最終行を取得

WSAO取込用.Activate
    For ループ = 1 To 最終行取得(WS計算, 1)

        WSAO取込用.Range(Cells(WSAO取込用最終行 + 1, 1), Cells(WSAO取込用最終行 + 1, 2)).Value = WSAO取込用.Cells(2, 1).Value
        WSAO取込用.Range(Cells(WSAO取込用最終行 + 2, 1), Cells(WSAO取込用最終行 + 2, 2)).Value = WSAO取込用.Cells(2, 1).Value

        WSAO取込用.Range(Cells(WSAO取込用最終行 + 1, 24), Cells(WSAO取込用最終行 + 1, 24)).Value = WSAO取込用.Cells(2, 1).Value
        WSAO取込用.Range(Cells(WSAO取込用最終行 + 2, 24), Cells(WSAO取込用最終行 + 2, 24)).Value = WSAO取込用.Cells(2, 1).Value

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 11).Formula = Format(3, "'00")
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 11).Formula = Format(4, "'00")

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 12).Value = WS計算.Cells(ループ, 2)
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 12).Value = WS計算.Cells(ループ, 2)

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 19).Value = WS計算.Cells(ループ, 3)
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 19).Value = WS計算.Cells(ループ, 4)

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 18).Value = WS計算.Cells(ループ, 6)
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 18).Value = WS計算.Cells(ループ, 5)

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 20).Value = _
                                                  WSAO取込用.Cells(WSAO取込用最終行 + 1, 18) * WSAO取込用.Cells(WSAO取込用最終行 + 1, 19)

        WSAO取込用.Cells(WSAO取込用最終行 + 2, 20).Value = _
                                                  WSAO取込用.Cells(WSAO取込用最終行 + 2, 18) * WSAO取込用.Cells(WSAO取込用最終行 + 2, 19)

        If InStr(ws貼付.Cells(1, 6).Value, "FDW") > 0 Then

            WSAO取込用.Cells(WSAO取込用最終行 + 1, 10).Formula = "FDW" & ws貼付.Cells(1, 7)
            WSAO取込用.Cells(WSAO取込用最終行 + 2, 10).Formula = "FDW" & ws貼付.Cells(1, 7)

        ElseIf InStr(ws貼付.Cells(1, 6).Value, "YDW") > 0 Then

            WSAO取込用.Cells(WSAO取込用最終行 + 1, 10).Formula = "YDW" & ws貼付.Cells(1, 7)
            WSAO取込用.Cells(WSAO取込用最終行 + 2, 10).Formula = "YDW" & ws貼付.Cells(1, 7)

        End If

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 5).Formula = Format(37001, "'00000000")
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 5).Formula = Format(37001, "'00000000")

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 7).Formula = Format(80000143, "'00000000")
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 7).Formula = Format(80000143, "'00000000")

        WSAO取込用.Cells(WSAO取込用最終行 + 1, 9).Formula = Format(998, "'00000000")
        WSAO取込用.Cells(WSAO取込用最終行 + 2, 9).Formula = Format(998, "'00000000")

        ' 最終行を手動で更新
        WSAO取込用最終行 = WSAO取込用最終行 + 2
    Next ループ
End Sub

Private Sub 取込用シート保存()

    Set WSAO取込用 = ThisWorkbook.Worksheets("AO取込用")
    Set ws貼付 = ThisWorkbook.Worksheets("貼付")

    With WorksheetFunction

        If .Sum(WSAO取込用.Range("T:T")) = -(ws貼付.Cells(1, 10)) * 1 Then
            MsgBox "取込用エクセルをデスクトップに保存しました"

            WSAO取込用.Copy

            Dim 新ワークブック As Workbook
            Dim ファイルパス As String
            Dim ファイル名 As String

            Set 新ワークブック = ActiveWorkbook

            ' デスクトップのパスを取得
            ファイルパス = Environ("USERPROFILE") & "\Desktop\"

            ' 保存するファイル名を指定
            ファイル名 = WSAO取込用.Cells(2, 10) & ".xlsx"

            ' 新しいブックをデスクトップに保存
            新ワークブック.SaveAs ファイルパス & ファイル名, FileFormat:=xlOpenXMLWorkbook

            ' 新しいブックを閉じる
            新ワークブック.Close SaveChanges:=False
        Else
            MsgBox "金額合っていません"
        End If

    End With

End Sub

Sub ☆実行()

    Call 自動更新停止(0)

    Call マスタシート再計算

    Call 貼付シート整形

    Call 計算シート作成

    Call 取込用シート作成
    
    Call 上書確認停止(0)

    Call 取込用シート保存
    
    Call 上書確認再開(0)

    Call 自動更新再開(0)

End Sub

Public Sub 九州開始()

Dim ws3 As Worksheet
     Set ws3 = ThisWorkbook.Worksheets("貼付")
     
    Dim ws5 As Worksheet
     Set ws5 = ThisWorkbook.Worksheets("九州貼付")
     
    Dim ws3_Acol_Lastrow As Long
     
      ws3_Acol_Lastrow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
      
      Dim ws5_Acol_Lastrow As Long
     
      ws5_Acol_Lastrow = ws5.Cells(ws5.Rows.Count, 1).End(xlUp).Row
      
      ws3.Range(ws3.Cells(3, 1), ws3.Cells(5000, 15)).Delete
      
      ws5.Range(ws5.Cells(3, "D"), ws5.Cells(ws5_Acol_Lastrow, "D")).Copy
      ws3.Cells(3, "A").PasteSpecial Paste:=xlPasteValues
      
      ws5.Range(ws5.Cells(3, "K"), ws5.Cells(ws5_Acol_Lastrow, "K")).Copy
      ws3.Cells(3, "C").PasteSpecial Paste:=xlPasteValues
      
       ws5.Range(ws5.Cells(3, "N"), ws5.Cells(ws5_Acol_Lastrow, "Q")).Copy
      ws3.Cells(3, "H").PasteSpecial Paste:=xlPasteValues
      
       ws3.Cells(1, "J").Value = Application.WorksheetFunction.Sum(ws3.Range("K:K"))
      
      Dim str As String
      Dim part1 As String ' 文字部分 (YDW)
    Dim part2 As String ' 数字部分 (0331)
    
    str = ws5.Range("G1").Value
    
    part1 = Left(str, 3)

       
        part2 = Mid(str, 4, 4)
        ws3.Range("F1").Value = part1
        ws3.Range("G1").Value = "'" & part2
'      Dim i As Long
'
'      For i = 3 To ws3_Acol_Lastrow
'
'      Next i

End Sub



