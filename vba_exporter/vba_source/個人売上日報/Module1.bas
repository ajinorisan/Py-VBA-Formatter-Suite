
Option Explicit

Private Const INDENTSPACENUM As Long = 4  '何タブか

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Writer:            Ditflame
'Date:              2019/02/21
'Desc:              Get Subroutine and Function names list. Output to Immediate pain (by Debug.print)
'Example of exec:   call ProcNamesToImmediatePain("Indent_ReFormatter")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ProcNamesToImmediatePain(SourceObjName As String)
    Dim buf As String
    Dim i As Long

    With ThisWorkbook.VBProject.VBComponents(SourceObjName).CodeModule
        For i = 1 To .CountOfLines
            If buf <> .ProcOfLine(i, 0) Then
                buf = .ProcOfLine(i, 0)
                Debug.Print buf
            End If
        Next i
    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Writer:            Ditflame
'Date:              2019/02/21
'Desc:              VBA Pritter(SourceAutoIndenter)
'Example of exec:   call SourceIndenter("Indent_ReFormatter")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SourceIndenter(SourceObjName As String)
    Dim codeTxt As String           'ソース出力用バッファ
    Dim CodeLine As String          'コードの取得行(※行末を_で複数行連結したものは1行とみなす)
    Dim CodeLine_Cache As String    '行末を_で複数行連結したものを1行として処理するためのワーク
    Dim CodeLine_Judge As String    'コメント部分を排除したソース整形判断用文字列
    Dim flg_MultiLine As Boolean    'T:行末を_で複数行連結したものの処理中 F:1行
    Dim i As Long
    Dim indLevel As Long            'インデントレベル(インデントの深さ)

    With Application.VBE.VBProjects(1).VBComponents(SourceObjName).CodeModule
        For i = 1 To .CountOfLines
            CodeLine = Trim(.Lines(i, 1))

            'ダブルクオートで囲った文字列部落とす
            CodeLine_Judge = CodeLine_Cache & DoubleQuateEject(CodeLine)

            If (0 = InStr(CodeLine_Judge, "'")) And (Right(CodeLine_Judge, 2) = " _") Then
                '複数行連結の"_"を削ってキャッシュに入れる
                CodeLine_Cache = Left(CodeLine_Judge, Len(CodeLine_Judge) - 1)

                '通常の場合と同様、インデントレベルに応じて出力、出力バッファにためる
                codeTxt = codeTxt & String(indLevel * INDENTSPACENUM, " ") & commentIndentSpaceAdd(CodeLine) & vbCrLf

                '"_"による複数行連結処理がはじまったので2行目以降インデントレベル上げ
                If Not flg_MultiLine Then
                    flg_MultiLine = True
                    indLevel = indLevel + 1
                End If
            Else
                CodeLine_Cache = ""

                If InStr(CodeLine_Judge, "'") = 0 Then
                    'コメントないので元のソース行で判断する
                    CodeLine_Judge = CodeLine
                Else
                    'コメント部分落としてTrimしたもので判断する
                    CodeLine_Judge = Trim(Left(CodeLine_Judge, InStr(CodeLine_Judge, "'") - 1))
                End If

                Select Case True
                '関数/Subルーチン宣言
                Case LeftCheck(CodeLine_Judge, "End Function")
                    indLevel = indLevel - 1
                Case LeftCheck(CodeLine_Judge, "End Sub")
                    indLevel = indLevel - 1

                'With
                Case LeftCheck(CodeLine_Judge, "End With")
                    indLevel = indLevel - 1

                'For
                Case LeftCheck(CodeLine_Judge, "Next ")
                    indLevel = indLevel - 1

                'Case
                Case LeftCheck(CodeLine_Judge, "Case ")
                    indLevel = indLevel - 1
                Case LeftCheck(CodeLine_Judge, "End Select")
                    indLevel = indLevel - 1

                'IF
                Case LeftCheck(CodeLine_Judge, "Else")
                    indLevel = indLevel - 1
                Case LeftCheck(CodeLine_Judge, "End If")
                    indLevel = indLevel - 1

                'Do...Loop
                Case LeftCheck(CodeLine_Judge, "Loop")
                    indLevel = indLevel - 1

                End Select

                If (InStr(Trim(CodeLine_Judge), " ") = 0) And (Right(CodeLine_Judge, 1) = ":") Then
                    'ラベル(行の中に空白がなく、最後がセミコロンで終わる ※例 hogehoge:)の場合はインデントなしで出力
                    codeTxt = codeTxt & CodeLine & vbCrLf
                Else
                    If Left(.Lines(i, 1), 1) = "'" Then
                        'Trim前の状態で、行頭からコメントの場合はインデントなしで出力
                        codeTxt = codeTxt & CodeLine & vbCrLf
                    Else
                        '通常の場合はインデントレベルに応じて出力、出力バッファにためる
                        codeTxt = codeTxt & String(indLevel * INDENTSPACENUM, " ") & commentIndentSpaceAdd(CodeLine) & vbCrLf
                    End If
                End If

                Select Case True
                '関数/Subルーチン宣言
                Case LeftCheck(CodeLine_Judge, "Public ")
                    indLevel = indLevel + 1
                Case LeftCheck(CodeLine_Judge, "Private ")
                    indLevel = indLevel + 1
                Case LeftCheck(CodeLine_Judge, "Function ")
                    indLevel = indLevel + 1
                Case LeftCheck(CodeLine_Judge, "Sub ")
                    indLevel = indLevel + 1

                'With
                Case LeftCheck(CodeLine_Judge, "With ")
                    indLevel = indLevel + 1

                'For
                Case LeftCheck(CodeLine_Judge, "For ")
                    indLevel = indLevel + 1

                'Case
                Case LeftCheck(CodeLine_Judge, "Case ")
                    indLevel = indLevel + 1
                Case LeftCheck(CodeLine_Judge, "Select Case ")
                    indLevel = indLevel + 1

                'IF
                Case LeftCheck(CodeLine_Judge, "If ") And (Right(CodeLine_Judge, 4) = "Then")
                    indLevel = indLevel + 1
                Case LeftCheck(CodeLine_Judge, "Else")
                    indLevel = indLevel + 1

                'Do...Loop
                Case LeftCheck(CodeLine_Judge, "Do ")
                    indLevel = indLevel + 1
                Case CodeLine = "Do"
                    indLevel = indLevel + 1

                End Select

                '"_"による複数行連結処理がおわったのでインデントレベル下げ
                If flg_MultiLine Then
                    flg_MultiLine = False
                    indLevel = indLevel - 1
                End If
            End If
        Next i
    End With

    With Application.VBE.VBProjects(1).VBComponents().Add(vbext_ct_StdModule)
        .Name = SourceObjName & "_ReIndent_" & Format(Now, "YYmmDD_HHMMSS")
        .CodeModule.AddFromString codeTxt
    End With
End Sub

'チェック文字列で切って文字列チェックする
Private Function LeftCheck(CodeLine As String, CheckTxt As String) As Boolean
    LeftCheck = (Left(CodeLine, Len(CheckTxt)) = CheckTxt)
End Function

'シングルクォートでのコメントを考慮しつつ、テキストリテラルを削る
Private Function DoubleQuateEject(CodeLine As String) As String
    Dim i As Long
    Dim s As String
    Dim isTxt As Boolean

    For i = 1 To Len(CodeLine)
        s = Mid(CodeLine, i, 1)

        If s = "'" And Not isTxt Then
            '行末までコメントなので全部返却して終了
            DoubleQuateEject = DoubleQuateEject & Mid(CodeLine, i)
            Exit Function
        End If

        If s = """" Then
            isTxt = Not isTxt
        Else
            If Not isTxt Then
                DoubleQuateEject = DoubleQuateEject & s
            End If
        End If
    Next i
End Function

'シングルクォート以降のコメントをインデント境界にあわせる
Private Function commentIndentSpaceAdd(In_Str As String) As String
    Dim SepStr1 As String
    Dim SepStr2 As String
    Dim SepLen As Long
    Dim HankakuLen As Long
    Dim AddSpaceLen As Long

    '1:テキストリテラルを加味してコメント開始位置を取得する
    SepLen = instrSingleQuoteWithOutText(In_Str)

    If SepLen = 0 Then
        'コメントないのでそのまま返却
        commentIndentSpaceAdd = In_Str
        Exit Function
    End If

    '2:コメント前後でテキストを分割する
    SepStr1 = Mid(In_Str, 1, SepLen - 1)
    SepStr2 = Mid(In_Str, SepLen)

    '3:全角半角を考慮して半角相当の文字数を出す
    HankakuLen = LenB(StrConv(SepStr1, vbFromUnicode))

    '4:半角スペース追加分の文字数を出す
    AddSpaceLen = (INDENTSPACENUM - (HankakuLen Mod INDENTSPACENUM)) Mod INDENTSPACENUM

    '5:組立てて完成
    commentIndentSpaceAdd = SepStr1 & String(AddSpaceLen, " ") & SepStr2

End Function

'テキストリテラルを加味してコメント開始位置を取得する
Private Function instrSingleQuoteWithOutText(In_Str As String) As Long
    Dim i As Long
    Dim s As String
    Dim isTxt As Boolean

    For i = 1 To Len(In_Str)
        s = Mid(In_Str, i, 1)

        If s = "'" And Not isTxt Then
            'ここからコメント
            instrSingleQuoteWithOutText = i
            Exit Function
        End If

        If s = """" Then
            isTxt = Not isTxt
        End If
    Next i
End Function

