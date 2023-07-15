Attribute VB_Name = "Module1"
Sub ¥¨¶°1()
Attribute ¥¨¶°1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ¥¨¶°1 ¥¨¶°
'

'
    ActiveWorkbook.Worksheets("¤u§@ªí1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("¤u§@ªí1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("¤u§@ªí1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub ¥¨¶°2()
Attribute ¥¨¶°2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ¥¨¶°2 ¥¨¶°
'

'
    ActiveWorkbook.Worksheets("¤u§@ªí1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("¤u§@ªí1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("¤u§@ªí1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E12").Select
End Sub
