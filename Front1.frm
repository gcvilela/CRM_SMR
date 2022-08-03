VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Front1 
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   OleObjectBlob   =   "Front1.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Front1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub c_Click()

End Sub

Private Sub cbbestado_Change()

End Sub

Private Sub cbbperfil_Change()



End Sub

Private Sub cblimpar_Click()

ListBox1.Enabled = False
Front1.ListBox1.Clear
Call clear_cadastro

maioridade.Value = True
txtpais.Value = "BRASIL"
cbbestado = "MG"

End Sub

Private Sub cbname1_AfterUpdate()
If cout = False Then
    Call Nomeapenas
Else
    Call BuscarBdNome
End If
End Sub

Private Sub cbname1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub cbname1_Change()

End Sub

Private Sub cbname1_Click()

End Sub

Private Sub cbname1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub cbsalvar_Click()
ListBox1.Enabled = False
'If IsDate(txtdata.Value) = False Then
'    MsgBox "Data Inválida!", vbCritical, "Atenção"
'    txtdata.BackColor = &HC0C0FF
'    txtdata.Value = ""
'    txtdata.SetFocus
'    Exit Sub
'End If
'
'If btsim.Value = True Then
'    If IsDate(txtmudança.Value) = False Then
'        MsgBox "Data Inválida!", vbCritical, "Atenção"
'        txtmudança.BackColor = &HC0C0FF
'        txtmudança.Value = ""
'        txtmudança.SetFocus
'        Exit Sub
'    End If
If campos(11) = "" Then
    MsgBox "Perfil em Branco!", vbCritical, "Atenção"
    campos(11).SetFocus
    Exit Sub
End If
If campos(1) = "" Then
    MsgBox "Nome em Branco!", vbCritical, "Atenção"
    campos(1).SetFocus
    Exit Sub
Else:
    Plan1Cb.Cells(hist, 16) = txtnome.Value
    Plan1Cb.Cells(hist, 15) = "CADASTRO      "
    hist = hist + 1
    Plan1Cb.Cells.Range("Q4") = hist
    
    If g = True Then
    
        Call update_cad
        Front1.ListBox1.Clear
        Call clear_cadastro
        maioridade.Value = True
        txtpais.Value = "BRASIL"
        cbbestado = "MG"
        
    Else:
    
        Call PreencherBase
        Front1.ListBox1.Clear
        Call clear_cadastro
        
        maioridade.Value = True
        txtpais.Value = "BRASIL"
        cbbestado = "MG"
    End If
End If


End Sub


Private Sub Frame1_Click()

End Sub

Private Sub ComboBox4_Change()

End Sub

Private Sub CommandButton14_Click()
list123.Clear
TextBox27.Value = ""
End Sub

Private Sub CommandButton15_Click()
Dim i As Integer
For i = 48 To 54
    If i <> 51 Then
        If campos(i) = "" Then
            MsgBox "Dado em Branco!", vbCritical, "Atenção"
            campos(i).SetFocus
            Exit Sub
        End If
    End If

Next
For i = 56 To 56
If OptionButton1.Value = False Then
    If campos(i) = "" Then
        MsgBox "Dado em Branco!", vbCritical, "Atenção"
        campos(i).SetFocus
        Exit Sub
    End If
End If
Next
For i = 57 To 57

    If campos(i) = "" Then
        MsgBox "Dado em Branco!", vbCritical, "Atenção"
        campos(i).SetFocus
        Exit Sub
    End If

Next
For i = 59 To 59

    If campos(i) = "" Then
        MsgBox "Dado em Branco!", vbCritical, "Atenção"
        campos(i).SetFocus
        Exit Sub
    End If

Next

    If campos(61) = "" Then
        MsgBox "Dado em Branco!", vbCritical, "Atenção"
        campos(61).SetFocus
        Exit Sub
    End If

Plan1Cb.Cells(hist, 16) = cbname1.Value
Plan1Cb.Cells(hist, 15) = "ATENDIMENTO"
hist = hist + 1
Plan1Cb.Cells.Range("Q4") = hist

If Front1.CheckBox1 = True Then
        Call atual_status
        
End If

If g = True Then
    Call update_cad2
Else:
    Call PreencherBase2
End If



CommandButton22.Value = True
ToggleButton1.Value = True
cout = False

For i = 48 To 61
    If i = 55 Then
        campos(55) = True
    Else:
        campos(i) = ""
    End If
Next

Plan1Cb.Activate
    'Dia da semana
    diasemanacb.Value = Plan1Cb.Cells(4, 9).Value
    
    'Data de hoje
    datatxt.Value = Plan1Cb.Cells(2, 9).Value
ComboBox6.Value = "CONSULTA"
ComboBox10.Value = "EM TRATAMENTO"
listboxtodos.Clear

Front1.cbname1.Value = ""
listboxtodos.Clear
End Sub

Private Sub CommandButton16_Click()
'CommandButton22.Value = True
Call Nomeapenas

End Sub

Private Sub CommandButton17_Click()

End Sub

Private Sub CommandButton18_Click()
Call pesquisar
End Sub

Private Sub CommandButton19_Click()
Dim i As Integer


CommandButton25.Visible = False

If campos(1) = "" Then
    MsgBox "Nome em Branco!", vbCritical, "Atenção"
    campos(1).SetFocus
    Exit Sub
End If
For i = 34 To 47
    campos(i) = ""
Next

MultiPage1.Value = 3
Frame21.Enabled = True

CommandButton25.Visible = False


End Sub

Private Sub CommandButton20_Click()



If campos(1) = "" Then
    MsgBox "Nome em Branco!", vbCritical, "Atenção"
    MultiPage1.Value = 0
    campos(1).SetFocus
    Exit Sub
End If

MultiPage1.Value = 0
Call preencheratendimento






End Sub

Private Sub CommandButton22_Click()
Dim i As Integer
CommandButton16.Visible = True
ToggleButton1.Value = True
cout = False

For i = 48 To 61
    If i = 55 Then
        campos(55) = True
    Else:
        campos(i) = ""
    End If
Next

Plan1Cb.Activate
    'Dia da semana
    diasemanacb.Value = Plan1Cb.Cells(4, 9).Value
    
    'Data de hoje
    datatxt.Value = Plan1Cb.Cells(2, 9).Value
Front1.ComboBox10 = "EM TRATAMENTO"
ComboBox6.Value = "CONSULTA"
listboxtodos.Clear

cbsemanames.Value = Plan1Cb.Cells.Range("Q5")
CommandButton26.Visible = False
g = False
End Sub

Private Sub CommandButton24_Click()
Dim i As Integer
CommandButton25.Visible = False
MultiPage1.Value = 0
Frame21.Enabled = False
For i = 34 To 47
    campos(i) = ""
Next
End Sub

Private Sub CommandButton25_Click()
Dim i As Integer

Call upadate_especial

CommandButton25.Visible = False

Frame21.Enabled = False

For i = 34 To 47
    campos(i) = ""
Next

Call add_list
End Sub

Private Sub CommandButton26_Click()
Dim i As Integer
            Call deletar_atendimento
            MsgBox "Atendimento excluído com sucesso!", vbInformation, "Informação"
            CommandButton22.Value = True
            CommandButton26.Visible = True
            Front1.cbname1.Value = id
            Call BuscarBdNome
            MultiPage1.Value = 1
            For i = 48 To 61
                If i = 55 Then
                    campos(55) = True
                Else:
                    campos(i) = ""
                End If
            Next

End Sub

Private Sub CommandButton27_Click()
Plan1Cb.Activate
Plan1Cb.Cells.Range("R2") = "Ordenado"
Call historic

End Sub

Private Sub datatxt_Change()

Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = datatxt.Value
datatxt.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
datatxt.Value = data3

End Sub

Private Sub datatxt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Frame10_Click()

End Sub

Private Sub Frame12_Click()

End Sub

Private Sub Frame13_Click()

End Sub

Private Sub Frame16_Click()

End Sub

Private Sub Frame19_Click()

End Sub

Private Sub Frame24_Click()

End Sub

Private Sub Frame5_Click()

End Sub

Private Sub Frame9_Click()

End Sub

Private Sub horatxt_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = horatxt.Value
horatxt.MaxLength = 5
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & ":" & Right(data3, 1)
        End If
        Next
    
horatxt.Value = data3
End Sub

Private Sub Label102_Click()

End Sub

Private Sub Label105_Click()

End Sub

Private Sub Label109_Click()

End Sub

Private Sub Label119_Click()

End Sub

Private Sub Label144_Click()

End Sub

Private Sub Label145_Click()

End Sub

Private Sub Label147_Click()

End Sub

Private Sub Label148_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label56_Click()

End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label68_Click()

End Sub

Private Sub Label74_Click()

End Sub

Private Sub list123_Click()


End Sub

Private Sub list123_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim n As Integer, i As Integer



Front1.ListBox1.Clear
n = list123.ListIndex
If n < 0 Then
    Exit Sub
End If
id = index1(n)

g = True

If OptionButton6 = True Then
    ListBox1.Enabled = True
    Call update_cadastro
    MultiPage1.Value = 0
    Call add_list
    list123.Clear
    Exit Sub
End If
If OptionButton5 = True Then
    CommandButton26.Visible = True
    Front1.cbname1.Value = id
    CommandButton16.Visible = False
    Call BuscarBdNome
    MultiPage1.Value = 1
    For i = 48 To 61
        If i = 55 Then
            campos(55) = True
        Else:
            campos(i) = ""
        End If
    Next
End If

list123.Clear

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim n As Integer

CommandButton25.Visible = True

n = ListBox1.ListIndex
If n < 0 Then
    Exit Sub
End If

Call Update_especialidade

MultiPage1.Value = 3
Frame21.Enabled = True

id_status = cbbstatus.Value
id_Profissional = cbbproficional.Value

End Sub

Private Sub listboxtodos_Click()

End Sub

Private Sub listboxtodos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim n As Integer

If g = False Then
        n = listboxtodos.ListIndex
        If n < 0 Then
        Exit Sub
        End If
        Front1.cbname1.Value = index1(n)
Else:
        n = listboxtodos.ListIndex
        If n < 0 Then
        Exit Sub
        End If
        
        'Call dataid
        'data_id = index1(n)
        
        'Call especialid
        'especial_id = index1(n)
        
        Call update_atendimento
End If
End Sub

Private Sub maioridade_Change()
Dim i As Integer
If maioridade.Value = True Then
For i = 17 To 32
    campos(i) = ""
Next
Frame22.Enabled = False
Else
Frame22.Enabled = True
End If

End Sub

Private Sub maioridade_Click()

End Sub

Private Sub OptionButton1_Click()
ComboBox4.Value = ""
End Sub


Private Sub OptionButton2_Change()

End Sub

Private Sub OptionButton2_Click()
If maioridade.Value = True Then
ComboBox4.Enabled = True
ComboBox4.Value = "NÃO JUSTIFICADO"
Else
ComboBox4.Enabled = False
End If

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TextBox55_Change()

End Sub

Private Sub OptionButton6_Click()

End Sub

Private Sub OptionButton7_Click()
Dim i As Integer

For i = 4 To 10
    If campos(i + 16) = "" Then
        campos(i + 16) = campos(i)
    End If
Next
End Sub

Private Sub OptionButton8_Change()

End Sub

Private Sub OptionButton8_Click()

End Sub

Private Sub TextBox27_AfterUpdate()
Call pesquisar
End Sub

Private Sub TextBox27_Change()

End Sub

Private Sub TextBox80_Change()

End Sub

Private Sub TextBox79_Change()

End Sub

Private Sub TextBox86_Change()

End Sub

Private Sub TextBox27_Enter()

End Sub

Private Sub ToggleButton1_Click()


If ToggleButton1 = True Then
cout = False
'Else
'cout = True
End If

End Sub

Private Sub txtassinatura_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtassinatura.Value
txtassinatura.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtassinatura.Value = data3
End Sub

Private Sub txtassinatura_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtbairro_Change()

End Sub

Private Sub txtcep_Change()
Dim cep As String, cep2 As String, cep3 As String
Dim i As Integer, j As Integer, n As Integer

cep = txtcep.Value
txtcep.MaxLength = 9
i = Len(cep)

    For j = 1 To i
        If IsNumeric(Mid(cep, j, 1)) Then
            cep2 = cep2 & Mid(cep, j, 1)
        End If
    Next
i = Len(cep2)
    For j = 1 To i
        cep3 = cep3 & Mid(cep2, j, 1)
        If j = 6 Then
        n = Len(cep3) - 1
            cep3 = Left(cep3, n) & "-" & Right(cep3, 1)
        End If
        Next
    
txtcep.Value = cep3
End Sub

Private Sub txtcep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtcidade_Click()

End Sub

Private Sub txtcpf_Change()

Dim CPF As String, CPF2 As String, CPF3 As String
Dim i As Integer, j As Integer, n As Integer

CPF = txtcpf.Value
txtcpf.MaxLength = 14
i = Len(CPF)

    For j = 1 To i
        If IsNumeric(Mid(CPF, j, 1)) Then
            CPF2 = CPF2 & Mid(CPF, j, 1)
        End If
    Next
i = Len(CPF2)
    For j = 1 To i
        CPF3 = CPF3 & Mid(CPF2, j, 1)
        If j = 4 Or j = 7 Then
        n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "." & Right(CPF3, 1)
        ElseIf j = 10 Then
        n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "-" & Right(CPF3, 1)
        End If
        Next
    
txtcpf.Value = CPF3


End Sub

Private Sub txtcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub txtcpf1_Change()
Dim CPF As String, CPF2 As String, CPF3 As String
Dim i As Integer, j As Integer, n As Integer

CPF = txtcpf.Value
txtcpf.MaxLength = 14
i = Len(CPF)

    For j = 1 To i
        If IsNumeric(Mid(CPF, j, 1)) Then
            CPF2 = CPF2 & Mid(CPF, j, 1)
        End If
    Next
i = Len(CPF2)
    For j = 1 To i
        CPF3 = CPF3 & Mid(CPF2, j, 1)
        If j = 4 Or j = 7 Then
        n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "." & Right(CPF3, 1)
        ElseIf j = 10 Then
        n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "-" & Right(CPF3, 1)
        End If
        Next
    
txtcpf.Value = CPF3
End Sub

Private Sub txtcpf1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtdata_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtdata.Value
txtdata.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtdata.Value = data3
End Sub


Private Sub txtdata_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


Private Sub txtdata1_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtdata1.Value
txtdata1.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtdata1.Value = data3
End Sub

Private Sub txtdata1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtencerramento_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtencerramento.Value
txtencerramento.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtencerramento.Value = data3
End Sub

Private Sub txtencerramento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtinicio_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtinicio.Value
txtinicio.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtinicio.Value = data3
End Sub

Private Sub txtmudança_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtmudança.Value
txtmudança.MaxLength = 10
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & "/" & Right(data3, 1)
        End If
        Next
    
txtmudança.Value = data3
End Sub

Private Sub txtinicio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtnome_Change()

End Sub

Private Sub txtnumerosus_Change()

End Sub

Private Sub txtpais_Change()

End Sub

Private Sub txtprogamada_Change()
Dim data As String, data2 As String, data3 As String
Dim i As Integer, j As Integer, n As Integer

data = txtprogamada.Value
txtprogamada.MaxLength = 5
i = Len(data)

    For j = 1 To i
        If IsNumeric(Mid(data, j, 1)) Then
            data2 = data2 & Mid(data, j, 1)
        End If
    Next
i = Len(data2)
    For j = 1 To i
        data3 = data3 & Mid(data2, j, 1)
        If j = 3 Then
        n = Len(data3) - 1
            data3 = Left(data3, n) & ":" & Right(data3, 1)
        End If
        Next
    
txtprogamada.Value = data3
End Sub

Private Sub txtprogamada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtrg_Change()

End Sub

Private Sub txtrg_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

End Sub

Private Sub txttelefone1_Change()

End Sub

Private Sub txttelefone1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txttelefone2_Change()

End Sub

Private Sub txttelefone2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Call hist_e
hist = Plan1Cb.Cells.Range("Q4")
'InitMaxMin Me.Caption

MultiPage1.Value = 0

g = False
cout = False
CommandButton25.Visible = False
txtnome.SetFocus
Frame21.Enabled = False
Frame22.Enabled = False
ComboBox4.Enabled = False
ToggleButton1 = True
ListBox1.Enabled = False
CommandButton26.Visible = False

ComboBox6.Value = "CONSULTA"
txtpais.Value = "BRASIL"
cbbestado = "MG"
ComboBox10 = "EM TRATAMENTO"



Call AtribuiCampos
'Call BuscarBD
Call dia_hoje
cbsemanames.Value = Plan1Cb.Cells.Range("Q5")
Plan1Cb.Activate
    'Dia da semana
    diasemanacb.Value = Plan1Cb.Cells(4, 9).Value
    
    'Data de hoje
    datatxt.Value = Plan1Cb.Cells(2, 9).Value

    ' --- function Combobox ---
    
      Dim cntr As Integer, i As Integer, coluna As Integer, a As Integer
      '- Combobox semana do mes
        For i = 1 To 5
            cbsemanames.AddItem i & "º"
        Next i
        
      '- Combobox do perfil
        cntr = Application.WorksheetFunction.CountA(Range("A:A"))
        For i = 2 To cntr
            coluna = 1
            cbbperfil.AddItem Plan1Cb.Cells(i, coluna)
            cbbperfil1.AddItem Plan1Cb.Cells(i, coluna)
        Next i

      '- Combobox do estado
        cntr = Application.WorksheetFunction.CountA(Range("B:B"))
        For i = 2 To cntr
            coluna = 2
            cbbestado.AddItem Plan1Cb.Cells(i, coluna)
            cbbestado1.AddItem Plan1Cb.Cells(i, coluna)
        Next i

      '- Combobox do tipo do plano
        cntr = Application.WorksheetFunction.CountA(Range("c:c"))
        For i = 2 To cntr
            coluna = 3
            cbbtipoplano.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
      '- Combobox da empresa do titular
        cntr = Application.WorksheetFunction.CountA(Range("d:d"))
        For i = 2 To cntr
            coluna = 4
            cbbempresa.AddItem Plan1Cb.Cells(i, coluna)
        Next i

      '- Combobox especialidade
        cntr = Application.WorksheetFunction.CountA(Range("e:e"))
        For i = 2 To cntr
            coluna = 5
            cbbespecialidade.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
         '- Combobox especialidade
        cntr = Application.WorksheetFunction.CountA(Range("e:e"))
        For i = 2 To cntr
            coluna = 5
            ComboBox5.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        

      '- Combobox Proficiaonal
        cntr = Application.WorksheetFunction.CountA(Range("f:f"))
        For i = 2 To cntr
            coluna = 6
            cbbproficional.AddItem Plan1Cb.Cells(i, coluna)
        Next i

      '- Combobox semana prog
        cntr = Application.WorksheetFunction.CountA(Range("j:j"))
        For i = 1 To cntr
            coluna = 10
            cbbprog.AddItem Plan1Cb.Cells(i, coluna)
            diasemanacb.AddItem Plan1Cb.Cells(i, coluna)
        Next i

      '- Combobox Perioticidade
        cntr = Application.WorksheetFunction.CountA(Range("g:g"))
        For i = 2 To cntr
            coluna = 7
            cbbperio.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
      '- Combobox status
        cntr = Application.WorksheetFunction.CountA(Range("H:H"))
        For i = 2 To cntr
            coluna = 8
            cbbstatus.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
        For i = 2 To cntr
            coluna = 8
            ComboBox10.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
      '- Combobox Ausente
        cntr = Application.WorksheetFunction.CountA(Range("k:k"))
        For i = 2 To cntr
            coluna = 11
            ComboBox4.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
        '- Combobox consulta
        cntr = Application.WorksheetFunction.CountA(Range("l:l"))
        For i = 2 To cntr
            coluna = 12
            ComboBox6.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        
        '- Combobox guia assinada
        cntr = Application.WorksheetFunction.CountA(Range("n:n"))
        
        For i = 2 To cntr
            coluna = 14
            TextBox74.AddItem Plan1Cb.Cells(i, coluna)
            ComboBox8.AddItem Plan1Cb.Cells(i, coluna)
        Next i
         For i = 2 To cntr - 1
            ComboBox9.AddItem Plan1Cb.Cells(i, coluna)
         Next
        
        '- Combobox consulta
        cntr = Application.WorksheetFunction.CountA(Range("m:m"))
        For i = 2 To cntr
            coluna = 13
            Textcidade.AddItem Plan1Cb.Cells(i, coluna)
        Next i
        

'Textcidade



    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim answer As Integer, ok As Integer
        answer = MsgBox("Deseja fechar o progama?", vbQuestion + vbYesNo + vbDefaultButton2, "")
        Application.ThisWorkbook.Save

        If CloseMode <> 1 Then
            If answer = vbYes Then
                'MsgBox "Yes"
            If Workbooks.Count = 1 Then
               Application.Visible = True
               Application.Quit
            Else
               Application.Visible = True
               ThisWorkbook.Close True
            End If

                Cancel = 0
            Else:
                'MsgBox "No"
                Cancel = 1
            End If

        End If


End Sub
