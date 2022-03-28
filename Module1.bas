Attribute VB_Name = "Module1"
Sub 口罩特約藥局排序()
Attribute 口罩特約藥局排序.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 口罩特約藥局排序 巨集
'
' 快速鍵: Ctrl+q
'
    Range("B1").Select
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
