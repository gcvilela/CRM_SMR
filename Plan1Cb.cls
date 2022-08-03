VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1Cb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mudar()

If Cells.Range("Q5") = "nao ordenado" Then
    Cells.Range("Q5") = "ordenado"
Else:
    Cells.Range("Q5") = "nao ordenado"
End If

End Sub
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
Application.EnableEvents = False
    If Target.Column = 18 Then
        Dim lastRow As Long
        lastRow = Cells(Rows.Count, 16).End(xlUp).Row
        Range("O2:P" & lastRow).Sort key1:=Range("P2:P" & lastRow), order1:=xlAscending, Header:=xlNo
    End If
Application.EnableEvents = True
    If Target.Column = 24 Then
            Plan1Cb.Activate
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
    End If
End Sub
