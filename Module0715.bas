Attribute VB_Name = "Module1"
Sub 巨集1()
Attribute 巨集1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 巨集1 巨集
'

'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 巨集2()
Attribute 巨集2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 巨集2 巨集
'

'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E12").Select
End Sub
