Sub Macro1()
    '
    ' Macro1 Macro
    '

    '
    ActiveSheet.Range("$A$1:$W$1664").AutoFilter Field:=2, Criteria1:=Array("-" _
    , "CS-A", "CS-B", "CS-C", "CS-D", "CS-E-IV", "CS-F-IV", "KT-KAO-GYU-01", _
    "KT-KAO-GYU-02", "KT-KAO-GYU-03", "KT-KAO-GYU-04", "WEB-GUTE-BL", "WEB-GUTE-BLL", _
    "WEB-GUTE-BM", "WEB-GUTE-BS", "YPC06-FR2-01", "YPC06-FR2-02", "YPC30-FR2-01", _
    "YPC30-FR2-02", "品番"), Operator:=xlFilterValues
    ActiveWindow.SmallScroll Down:=-12
    Sheets("集計").Select
End Sub
Sub Macro2()
    '
    ' Macro2 Macro
    '

    '
    ActiveSheet.Range("$A$1:$W$1664").AutoFilter Field:=2, Criteria1:="<>0", _
    Operator:=xlAnd
End Sub
Sub Macro3()
    '
    ' Macro3 Macro
    '

    '
    ActiveSheet.Range("$A$1:$S$1664").AutoFilter Field:=2, Criteria1:=Array("-" _
    , "CS-A", "CS-B", "CS-C", "CS-D", "CS-E-IV", "CS-F-IV", "KT-KAO-GYU-01", _
    "KT-KAO-GYU-02", "KT-KAO-GYU-03", "KT-KAO-GYU-04", "WEB-GUTE-BL", "WEB-GUTE-BLL", _
    "WEB-GUTE-BM", "WEB-GUTE-BS", "YPC06-FR2-01", "YPC06-FR2-02", "YPC30-FR2-01", _
    "YPC30-FR2-02", "品番"), Operator:=xlFilterValues
End Sub