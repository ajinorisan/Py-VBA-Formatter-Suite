Option Explicit

Sub スタイル削除()
    'スタイルを全削除
    On Error Resume Next

    Dim i
    Dim s

    'Indexを指定して全スタイル削除
    For i = ActiveWorkbook.Styles.Count To 1 Step -1
        s = ActiveWorkbook.Styles(i).Delete
    Next

End Sub