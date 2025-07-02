Option Explicit

Sub AO用データ作成()

    Application.ScreenUpdating = False

    'Dim WS3 As Worksheet
    'Set WS3 = ThisWorkbook.Worksheets("アラジン取込用(受注)")

    Dim WS4 As Worksheet
    Set WS4 = ThisWorkbook.Worksheets("アラジン取込用(売上)")

    Dim WS10 As Worksheet
    Set WS10 = ThisWorkbook.Worksheets("在庫用")

    Dim WS11 As Worksheet
    Set WS11 = ThisWorkbook.Worksheets("配送便確認用")

    'Dim WS3_AC_LR As Long
    'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row
    'Dim WS4_AC_LR As Long
    'WS4_AC_LR = WS4.Cells(WS4.Rows.Count, 1).End(xlUp).Row
    '
    'If (WS3_AC_LR <> WS4_AC_LR) And (WS3.Cells(2, 9) <> WS4.Cells(2, 10)) Then
    'WS3.Copy

    'ActiveWorkbook.Worksheets("アラジン取込用(受注)").OLEObjects.Delete
    '
    'WS3_AC_LR = WS3.Cells(WS3.Rows.Count, 1).End(xlUp).Row

    'Dim FOR1 As Long
    'For FOR1 = 2 To WS3_AC_LR
    'If ActiveWorkbook.Worksheets("アラジン取込用(受注)").Cells(FOR1, 10) = "02" Then
    'ActiveWorkbook.Worksheets("アラジン取込用(受注)").Cells(FOR1, 10).Formula = "'01"
    'End If
    'Next FOR1

    'End If
    WS10.Activate

    WS4.Range("AA:AD").Delete
    WS4.Copy
    WS11.Copy

    Application.ScreenUpdating = True

End Sub