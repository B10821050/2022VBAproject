Attribute VB_Name = "Module1"
Sub �f�n�S���ħ��Ƨ�()
Attribute �f�n�S���ħ��Ƨ�.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �f�n�S���ħ��Ƨ� ����
'
' �ֳt��: Ctrl+q
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
