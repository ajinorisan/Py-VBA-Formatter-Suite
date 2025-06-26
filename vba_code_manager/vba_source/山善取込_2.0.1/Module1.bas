Option Explicit

Sub 山善取込()
Dim WS1 As Worksheet
Set WS1 = ThisWorkbook.Worksheets("取込変換")
Dim WS2 As Worksheet
Set WS2 = ThisWorkbook.Worksheets("得意先納品先マスタ変換")
Dim WS3 As Worksheet
Set WS3 = ThisWorkbook.Worksheets("商品マスタ変換")

Application.ScreenUpdating = False

With WS1.Range("A:W")
        .NumberFormatLocal = "@"
        End With

'カレントフォルダを固定
 With CreateObject("WScript.Shell")
        .CurrentDirectory = "\\192.168.1.9\業務\　HC･雑貨･ｽﾃｰｼｮﾅﾘｰ\☆ 得意先情報\000037　山善\YES取り込みデータ"
    End With
    
    '取込変換シートの前回のデータを削除
    WS1.Rows("2:3000").Delete
    
    '山善EDI csvデータを開ける
    Dim OFN As String
     Dim FN As String
     Dim FL As Long
    OFN = Application.GetOpenFilename("山善データ,*.csv?")
If OFN <> "False" Then
 
    FN = Dir(OFN)
     FL = Len(FN)
    Else
    Application.ScreenUpdating = True
    Exit Sub
    End If

    Workbooks.Open FN
    
'    Dim AWB2 As Workbook
'   Set AWB2 = ActiveWorkbook
    
Dim AWS1 As Worksheet
Set AWS1 = ActiveWorkbook.Worksheets(1)
Dim AWS1_AC_LR As Long
AWS1_AC_LR = AWS1.Cells(AWS1.Rows.Count, 1).End(xlUp).Row
'AWS1.Columns("AQ:AQ").NumberFormatLocal = "0_ "

With WorksheetFunction

'データ転記
Dim FOR1 As Long
For FOR1 = 2 To AWS1_AC_LR

AWS1.Cells(FOR1, 43).Value = Format(AWS1.Cells(FOR1, 43), "0000000000000")

WS1.Cells(FOR1, 1).Value = Format(AWS1.Cells(FOR1, 3), "yyyymmdd")

On Error Resume Next
WS1.Cells(FOR1, 6).Value = _
.Index(WS2.Cells, .Match(Mid(AWS1.Cells(FOR1, 7), 1, 3) & "-" & AWS1.Cells(FOR1, 30), WS2.Range("D:D"), 0), 7)

WS1.Cells(FOR1, 4).Value = _
.Index(WS2.Cells, .Match(Mid(AWS1.Cells(FOR1, 7), 1, 3) & "-" & AWS1.Cells(FOR1, 30), WS2.Range("D:D"), 0), 6)

WS1.Cells(FOR1, 11).Value = _
.Index(WS3.Cells, .Match(AWS1.Cells(FOR1, 43) * 1, WS3.Range("E:E"), 0), 2)

WS1.Cells(FOR1, 18).Value = _
.Index(WS3.Cells, .Match(AWS1.Cells(FOR1, 43) * 1, WS3.Range("E:E"), 0), 4)
On Error GoTo 0

If WS1.Cells(FOR1, 18) = 0 Then
WS1.Cells(FOR1, 18).Value = _
AWS1.Cells(FOR1, 47)
End If

WS1.Cells(FOR1, 17).Value = _
AWS1.Cells(FOR1, 44)

WS1.Cells(FOR1, 22).Value = _
AWS1.Cells(FOR1, 7)

WS1.Cells(FOR1, 9).Value = _
AWS1.Cells(FOR1, 7)

WS1.Cells(FOR1, 8).Value = _
"'00000001"

WS1.Cells(FOR1, 10).Value = _
"'01"

WS1.Cells(FOR1, 2).Value = _
AWS1.Cells(FOR1, 48)

Next FOR1
    
    '例外納品先処理及びコード登録無しの商品処理
For FOR1 = 2 To AWS1_AC_LR
If WS1.Cells(FOR1, 6) = "" Then
WS1.Cells(FOR1, 6).Value = "80013868"
End If
If WS1.Cells(FOR1, 11) = "" Then
    WS1.Cells(FOR1, 11).Value = "Z-00011-1"
    WS1.Cells(FOR1, 14).Value = "山善CD:" & AWS1.Cells(FOR1, 39) & " " & AWS1.Cells(FOR1, 43)
    WS1.Cells(FOR1, 13).Value = Mid(AWS1.Cells(FOR1, 42), 1, 38)
    End If
    
  Next FOR1
    End With
    
    '得意先マスタ登録なし処理
For FOR1 = 2 To AWS1_AC_LR
If WS1.Cells(FOR1, 4) = "" Then
    GoTo ERR1
    End If
    
    Next FOR1
    
   '取込用エクセルーデータはき出し処理
    WS1.Copy
   
   Dim AWB As Workbook
   Set AWB = ActiveWorkbook
   Dim AWBWS1 As Worksheet
   Set AWBWS1 = AWB.Worksheets("取込変換")
   Dim AWBWS1_AC_LR As Long
   AWBWS1_AC_LR = AWBWS1.Cells(AWBWS1.Rows.Count, 1).End(xlUp).Row
   AWBWS1.Rows(AWBWS1_AC_LR + 1 & ":10000").Delete
   AWBWS1.OLEObjects.Delete
   
    Application.DisplayAlerts = False
'    On Error GoTo MYERROR

Dim WSH As Object
  Set WSH = CreateObject("WScript.Shell")
' wsh.SpecialFolders ("Desktop")
  
    AWB.SaveAs Filename:=WSH.SpecialFolders("Desktop") & "\" & Format(Now, "yyyymmddhhmmss") & "山善取込.xlsx"
    Set WSH = Nothing
     
    AWB.Close
'AWB2.Close
     
    Workbooks(FN).Close
    Application.DisplayAlerts = True
    If CurDir <> "\\192.168.1.9\業務\　HC･雑貨･ｽﾃｰｼｮﾅﾘｰ\☆ 得意先情報\000037　山善\YES取り込みデータ\取込済" Then
    FileCopy "\\192.168.1.9\業務\　HC･雑貨･ｽﾃｰｼｮﾅﾘｰ\☆ 得意先情報\000037　山善\YES取り込みデータ\" & Mid(FN, 1, FL - 4) & ".csv" _
    , "\\192.168.1.9\業務\　HC･雑貨･ｽﾃｰｼｮﾅﾘｰ\☆ 得意先情報\000037　山善\YES取り込みデータ\取込済\" & Mid(FN, 1, FL - 4) & ".csv"
   Kill "\\192.168.1.9\業務\　HC･雑貨･ｽﾃｰｼｮﾅﾘｰ\☆ 得意先情報\000037　山善\YES取り込みデータ\" & Mid(FN, 1, FL - 4) & ".csv"
   End If

    Application.ScreenUpdating = True
    MsgBox "ファイル出力完了"
    Exit Sub
'ER1:
'    MsgBox "ファイルの場所が不正です。"
'    Application.ScreenUpdating = True
'    Exit Sub
ERR1:
Application.DisplayAlerts = False
Workbooks(FN).Close
Application.DisplayAlerts = True
Application.ScreenUpdating = True

    MsgBox FOR1 & "行の得意先を得意先納品先マスタ変換に登録してください"
    
    End Sub


