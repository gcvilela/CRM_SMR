Attribute VB_Name = "Módulo1"

Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("Tabela1[#All]").Select
    ActiveWorkbook.Worksheets("Combobox1").ListObjects("Tabela1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Combobox1").ListObjects("Tabela1").Sort.SortFields. _
        Add2 Key:=Range("Tabela1[DATA]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, CustomOrder:="jan,fev,mar,abr,mai,jun,jul,ago,set,out,nov,dez" _
        , DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Combobox1").ListObjects("Tabela1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
