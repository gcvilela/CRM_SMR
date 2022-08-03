Attribute VB_Name = "Geral"
Option Explicit

Global campos(1 To 100) As Object
Global id As String, x_spl() As String
Global new_id As String, id_status As String, id_Profissional As String
Global especial_id As String, profi_atual As String
Global data_id As String, hist As Integer
Global g As Boolean, cout As Boolean
Global index1() As String

'Declare PtrSafe Function FindWindowA& Lib "User32" (ByVal lpClassName$, ByVal lpWindowName$)
'Declare PtrSafe Function GetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&)
'Declare PtrSafe Function SetWindowLongA& Lib "User32" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
Public Const GWL_STYLE As Long = -16
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000

'Public Sub InitMaxMin(mCaption As String, Optional Max As Boolean = True, Optional Min As Boolean = True
       ' , Optional Sizing As Boolean = False)
'Dim hWnd As Long
   ' hWnd = FindWindowA(vbNullString, mCaption)
   ' If Max Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MAXIMIZEBOX
   ' If Min Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_MINIMIZEBOX
   ' If Sizing Then SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) Or WS_FULLSIZING
    
    
    
'End Sub

Sub AtribuiCampos()

Dim i As Integer


    'for i = 1 to 33  Cadastro 33
    Set campos(1) = Front1.txtnome:         Set campos(2) = Front1.txtcpf:
    Set campos(3) = Front1.txtrg:           Set campos(4) = Front1.txtend:
    Set campos(5) = Front1.txtbairro:       Set campos(6) = Front1.Textcidade:
    Set campos(7) = Front1.cbbestado:       Set campos(8) = Front1.txtpais:
    Set campos(9) = Front1.txtcep:          Set campos(10) = Front1.txtcomp:
    Set campos(11) = Front1.cbbperfil:      Set campos(12) = Front1.txtdata:
    Set campos(13) = Front1.txttelefone1:   Set campos(14) = Front1.txttelefone2:
    Set campos(15) = Front1.txtemail:       Set campos(16) = Front1.txtobs:
'-------------------------------------------------------------------------
    Set campos(17) = Front1.txtnome1:       Set campos(18) = Front1.txtcpf1:
    Set campos(19) = Front1.txtrg1:         Set campos(20) = Front1.txtend1:
    Set campos(21) = Front1.txtbairro1:     Set campos(22) = Front1.Textcidade1:
    Set campos(23) = Front1.cbbestado1:     Set campos(24) = Front1.txtpais1:
    Set campos(25) = Front1.txtcep1:        Set campos(26) = Front1.txtcomp1:
    Set campos(27) = Front1.cbbperfil1:     Set campos(28) = Front1.txtdata1:
    Set campos(29) = Front1.txttelefone11:  Set campos(30) = Front1.txttelefone21:
    Set campos(31) = Front1.txtemail1:      Set campos(32) = Front1.txtobs1:
    Set campos(33) = Front1.maioridade:
'------------------------------------------------------------------------
    'for i = 34 to 47 Especialidade 14
    Set campos(34) = Front1.cbbespecialidade: Set campos(35) = Front1.cbbproficional:
    Set campos(36) = Front1.cbbtipoplano:     Set campos(37) = Front1.txtnumerosus:
    Set campos(38) = Front1.txtnumeroplano:   Set campos(39) = Front1.cbbempresa:
    Set campos(40) = Front1.txtinicio:        Set campos(41) = Front1.txtassinatura:
    Set campos(42) = Front1.cbbstatus:        Set campos(43) = Front1.txtencerramento:
    Set campos(44) = Front1.cbbperio:         Set campos(45) = Front1.cbbprog:
    Set campos(46) = Front1.txtprogamada:     Set campos(47) = Front1.obs3:
'------------------------------------------------------------------------------
   'For i = 48 To 61 Atendimento 14
    Set campos(48) = Front1.cbname1:     Set campos(49) = Front1.datatxt:
    Set campos(50) = Front1.horatxt:     Set campos(51) = Front1.txtsenha:
    Set campos(52) = Front1.ComboBox5:   Set campos(53) = Front1.diasemanacb:
    Set campos(54) = Front1.cbsemanames: Set campos(55) = Front1.OptionButton1:
    Set campos(56) = Front1.ComboBox4:   Set campos(57) = Front1.ComboBox6:
    Set campos(58) = Front1.obstxt:      Set campos(59) = Front1.TextBox74:
    Set campos(60) = Front1.ComboBox9:   Set campos(61) = Front1.ComboBox10:
        
    
    
    
'----------------------------------------------------------
    

g = False

For i = 1 To 33
    campos(i).TabIndex = i - 1
Next
For i = 34 To 47
    campos(i).TabIndex = i - 34
Next
For i = 48 To 61
    campos(i).TabIndex = i - 48
Next

End Sub

Sub preencheratendimento()

Dim CampoFinal As String
Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, i As Integer

arq = ThisWorkbook.Path & "\" & "BD.accdb;"
CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"


CampoFinal = ""
For i = 34 To 47
        CampoFinal = CampoFinal & "'" & campos(i) & "'"
        If i < 47 Then CampoFinal = CampoFinal & ","
Next
'

SQL = "Insert Into especialidade(nome,Especialidade,Profissional,[Tipo do Plano],[Num Carteira do SUS],[Num Carteira do Plano],[Empresa do titular],[Data Início do Tratamento],[Data Assinatura do Contrato],Status,[Data Encerramento],Periodicidade,[Dia da Semana Progamado],[Hora Progamada],Obs) "
SQL = SQL & "Values(" & "'" & campos(1) & "'," & CampoFinal & ")"




BD.Open CS
RS.Open SQL, BD
Front1.ListBox1.AddItem campos(34) & "     " & campos(35) & "     " & campos(36)
BD.Close

MsgBox "Especialidade Salva com sucesso!", vbInformation, "Informação"
For i = 34 To 47
    campos(i) = ""
Next
Front1.Frame21.Enabled = False
End Sub
Sub PreencherBase()
If g = True Then
Exit Sub
End If

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String
Dim i As Integer
Dim CampoFinal As String


arq = ThisWorkbook.Path & "\" & "BD.accdb;"
CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"


CampoFinal = ""
    For i = 1 To 33
            CampoFinal = CampoFinal & "'" & campos(i) & "'"
            If i < 33 Then CampoFinal = CampoFinal & ","
    Next

SQL = "Insert Into dadospessoais(Nome,CPF,RG,Endereço,Bairro,Cidade,Estado," _
& "País,CEP,Complemento,Perfil,DataNasc,Telefone1,Telefone2,Email,Obs1," _
& "Nome2,CPF2,RG2,Endereço2,Bairro2,Cidade2,[Estado 2]," _
& "País2,CEP2,Complemento2,Perfil2,DataNasc2,Telefone12,Telefone22,Email2,Obs2,maioridade) "

SQL = SQL & "Values(" & CampoFinal & ")"



BD.Open CS
RS.Open SQL, BD

BD.Close

MsgBox "Cadastro efetuado com sucesso!", vbInformation, "Informação"
End Sub
Sub PreencherBase2()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String
Dim CampoFinal As String, i As Integer


If Front1.OptionButton1 = True Then
x = "PRESENTE"
End If
If Front1.OptionButton1 = False Then
x = "AUSENTE" & " "
End If


arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

'(nome,data,hora,senha,Especialidade,dia_semana,semana_do_mes,presente,motivo_au,com/retor,obs,guia_assinada,guia_recebida)

CampoFinal = ""
    For i = 48 To 61
            If i = 55 Then
                CampoFinal = CampoFinal & "'" & x & "'"
            Else
                CampoFinal = CampoFinal & "'" & campos(i) & "'"
            End If
            If i < 61 Then
            CampoFinal = CampoFinal & ","
            End If
            
            
    Next

SQL = "Insert Into consultas(nome,data,hora,senha,Especialidade,[dia_semana],[semana_do_mes],presente,[motivo_au],[com_retor],obs,[guia_assinada],[guia_recebida],status) "
SQL = SQL & "Values(" & CampoFinal & ")"

BD.Open CS
RS.Open SQL, BD

BD.Close
MsgBox "Cadastro efetuado com sucesso!", vbInformation, "Informação"


End Sub
Sub Nomeapenas()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, name As String

arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"




x = "'%" & Front1.cbname1 & "%'"
SQL = "Select * from dadospessoais WHERE Nome like " & x & ""
SQL = SQL + " ORDER by nome ASC"

BD.Open CS
RS.Open SQL, BD
x = ""
Front1.listboxtodos.Clear
Do Until RS.EOF
    Front1.listboxtodos.AddItem RS!Nome
    x = x & RS.Fields(0) & ","
    RS.MoveNext
Loop


index1() = Split(x, ",")



RS.Close
BD.Close



End Sub
Sub BuscarBdNome()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, name As String, i As Integer

arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"
x = "'" & Front1.cbname1 & "'"
SQL = "Select * from consultas WHERE Nome like " & x & " order by data DESC"

BD.Open CS

RS.Open SQL, BD
i = 2
Front1.listboxtodos.Clear
Plan1Cb.Cells.Range("S2:v1000") = ""
Do Until RS.EOF
    Plan1Cb.Cells.Range("s" & i) = RS!data
    Plan1Cb.Cells.Range("t" & i) = RS!Nome
    Plan1Cb.Cells.Range("u" & i) = RS!Presente
    Plan1Cb.Cells.Range("v" & i) = RS!id
    'Front1.listboxtodos.AddItem RS!Nome & "     " & RS!Presente & "     " & RS!data & "     " & RS!semana_do_mes & " semana"
    RS.MoveNext
    i = i + 1
Loop
RS.Close
Plan1Cb.Cells.Range("x2") = i

For i = 2 To Plan1Cb.Cells.Range("x2").Value
    Front1.listboxtodos.AddItem Plan1Cb.Cells.Range("t" & i) & "             " & Plan1Cb.Cells.Range("s" & i) & "         " & Plan1Cb.Cells.Range("u" & i)

Next
BD.Close

End Sub
Sub BuscarBD()

'Dim BD As New ADODB.Connection
'Dim RS As New ADODB.Recordset
'Dim arq As String, CS As String, SQL As String
'
'
'arq = ThisWorkbook.Path & "\" & "BD.accdb;"
'
'CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
'& "Data Source=" & arq _
'& "jet OLEDB:Database Password= 123456;"
''& "Persist Security Info=False;"
'
'SQL = "Select Nome from dadospessoais order by Nome"
'WHERE Nome like Front1.cbname1.Value
'BD.Open CS
'RS.Open SQL, BD
'
'Front1.cbname1.Clear
'Do Until RS.EOF
'    Front1.cbname1.AddItem RS!nome
'    RS.MoveNext
'Loop
'
'
'RS.Close
'BD.Close

End Sub
Sub especialid()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer, f As String
Dim split_S() As String, base As String, q As Variant

arq = ThisWorkbook.Path & "\" & "BD.accdb;"
CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"



SQL = "Select * from consultas WHERE nome = '" & id & "'"
SQL = SQL & " order by data DESC"

BD.Open CS
RS.Open SQL, BD
x = ""
Do Until RS.EOF
    Set q = RS.Fields(4):
    
    If IsNull(q) Then
        f = ","
    Else:
        f = q
    End If
    
    x = x & f & ","
    
RS.MoveNext
Loop

index1() = Split(x, ",")
    
RS.Close
BD.Close

End Sub
Sub dataid()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer, f As String
Dim split_S() As String, base As String, q As Variant

arq = ThisWorkbook.Path & "\" & "BD.accdb;"
CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"



SQL = "Select * from consultas WHERE nome = '" & id & "'"
SQL = SQL & " order by data DESC"


BD.Open CS
RS.Open SQL, BD
x = ""

Do Until RS.EOF
    Set q = RS.Fields(1):
    
    If IsNull(q) Then
        f = ","
    Else:
        f = q
    End If
    
    x = x & f & ","
    
RS.MoveNext
Loop

index1() = Split(x, ",")
    
RS.Close
BD.Close

End Sub
Sub newid()
Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer, f As String
Dim split_S() As String, base As String, q As Variant

arq = ThisWorkbook.Path & "\" & "BD.accdb;"
CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"



SQL = "Select * from especialidade WHERE nome = '" & id & "'"
SQL = SQL + " ORDER BY Profissional ASC"


BD.Open CS
RS.Open SQL, BD
x = ""
Do Until RS.EOF
    Set q = RS.Fields(1):
    
    If IsNull(q) Then
        f = ","
    Else:
        f = q
    End If
    
    x = x & f & ","
    
RS.MoveNext
Loop

index1() = Split(x, ",")
    
RS.Close
BD.Close


End Sub
Sub deletar_atendimento()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, f As String, f_split() As String, i As Integer, n As Integer
arq = ThisWorkbook.Path & "\" & "BD.accdb;"
Dim q As Variant


CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

n = Front1.listboxtodos.ListIndex
If n < 0 Then
    Exit Sub
End If
SQL = "DELETE from consultas "
'WHERE id = " & f_split(0) & ""
SQL = SQL & "Where id = " & "" & Plan1Cb.Cells.Range("v" & n + 2) & "" & ""

'SQL = "SELECT * from consultas WHERE nome = " & "'" & id & "'" & ""
'SQL = SQL & " AND Especialidade = " & "'" & especial_id & "'" & ""
'SQL = SQL & " AND data = " & "'" & data_id & "'" & ""
'SQL = SQL & " ORDER BY id ASC"


BD.Open CS
'f = ""
'Do Until RS.EOF
'
'    f = f & RS.Fields(14) & ","
'    RS.MoveNext
'Loop
'f_split() = Split(f, ",")
'
'RS.Close



'nome = " & "'" & id & "'" & ""
'SQL = SQL & " AND Especialidade = " & "'" & especial_id & "'" & ""
'SQL = SQL & " AND data = " & "'" & data_id & "'" & ""


BD.Execute SQL
BD.Close


End Sub

Sub update_atendimento()


Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, f As String, i As Integer, n As Integer
arq = ThisWorkbook.Path & "\" & "BD.accdb;"
Dim q As Variant


CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

n = Front1.listboxtodos.ListIndex
If n < 0 Then
    Exit Sub
End If
SQL = "Select * from consultas WHERE id = " & "" & Plan1Cb.Cells.Range("v" & n + 2) & "" & "" 'nome = " & "'" & id & "'" & ""
'SQL = SQL & " AND Especialidade = " & "'" & especial_id & "'" & ""
'SQL = SQL & " AND data = " & "'" & data_id & "'" & ""


BD.Open CS
RS.Open SQL, BD


For i = 0 To 13
    
    Set q = RS.Fields(i):
    If IsNull(q) Then
        f = ""
    Else:
        f = q
    End If
If i + 48 = 55 Then
    If q = "AUSENTE " Then
        Front1.OptionButton2.Value = True
    Else:
        Front1.OptionButton1.Value = True
    End If
Else:
    campos(i + 48) = f
End If

Next

End Sub
Sub Update_especialidade()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, f As String, i As Integer, i2 As Integer, n As Integer
arq = ThisWorkbook.Path & "\" & "BD.accdb;"
Dim q As Variant


CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

f = "'" & new_id & "'"
    
n = Front1.ListBox1.ListIndex
If n < 0 Then
    Exit Sub
End If

SQL = "Select * from especialidade WHERE id = " & x_spl(n) & ""
SQL = SQL + " Order by id DESC"


BD.Open CS
RS.Open SQL, BD

If Front1.ListBox1.ListIndex > 0 Then

End If
For i = 1 To 14
    Set q = RS.Fields(i):
    If IsNull(q) Then
        f = ""
    Else:
        
        f = q
    End If
campos(i + 33) = f

Next


RS.Close
BD.Close



End Sub
Sub add_list()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String

arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"


x = "'" & id & "'"
SQL = "Select * from especialidade WHERE Nome = " & x & ""
SQL = SQL + " Order by id ASC"
BD.Open CS
RS.Open SQL, BD

Front1.ListBox1.Clear
x = ""
Do Until RS.EOF
    x = x & RS.Fields(15) & ","
    Front1.ListBox1.AddItem RS.Fields(1) & "     " & RS.Fields(2) & "     " & RS.Fields(3)
    RS.MoveNext
Loop
x_spl() = Split(x, ",")

RS.Close
BD.Close



End Sub
Sub update_cadastro()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, f As String, i As Integer
Dim q As Variant
arq = ThisWorkbook.Path & "\" & "BD.accdb;"



CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

f = "'" & id & "'"

SQL = "Select * from dadospessoais WHERE Nome = " & f & ""

BD.Open CS
RS.Open SQL, BD


If RS.Fields(32) = "NÃO" Then
    Front1.OptionButton7 = True
    Else
    Front1.maioridade = True
End If

For i = 0 To 32
Set q = RS.Fields(i):
If IsNull(q) Then
    f = ""
Else:
    
    f = q
End If

campos(i + 1) = f

Next

RS.Close
BD.Close

End Sub
Sub update_cad2()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer, n As Integer
Dim split_S() As String, base As String


arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

base = "nome,data,hora,senha,Especialidade,dia_semana,semana_do_mes,presente,motivo_au,com_retor,obs,guia_assinada,guia_recebida,status"
split_S() = Split(base, ",")



SQL = "UPDATE consultas SET "
For i = 0 To 13
        If i = 7 Then
            If campos(i + 48) = True Then
                SQL = SQL & split_S(i) & " = 'PRESENTE"
                Else
                SQL = SQL & split_S(i) & " = 'AUSENTE "
            End If
        Else
            SQL = SQL & split_S(i) & " = '" & campos(i + 48)
        End If
        If i < 13 Then
        SQL = SQL & "',"
        Else:
        SQL = SQL & "' "
        End If
Next
n = Front1.listboxtodos.ListIndex
If n < 0 Then
    MsgBox "Não é possivel alterar esse campo!", vbInformation, "Informação"
    Exit Sub
End If

SQL = SQL & "Where id = " & "" & Plan1Cb.Cells.Range("v" & n + 2) & "" & ""
'SQL = SQL & "Where nome = " & "'" & id & "'"
'SQL = SQL & " AND Especialidade = " & "'" & especial_id & "'" & ""
'SQL = SQL & " AND data = " & "'" & data_id & "'" & ""



BD.Open CS
BD.Execute SQL
BD.Close

MsgBox "Alteração efetuada com sucesso!", vbInformation, "Informação"

End Sub
Sub upadate_especial()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer
Dim split_S() As String, base As String, n As Integer


arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

base = "Especialidade,Profissional,[Tipo do Plano],[Num Carteira do SUS]," _
& "[Num Carteira do Plano],[Empresa do titular],[Data Início do Tratamento]," _
& "[Data Assinatura do Contrato],Status,[Data Encerramento],Periodicidade," _
& "[Dia da Semana Progamado],[Hora Progamada],Obs"


split_S() = Split(base, ",")

SQL = "UPDATE especialidade SET "
For i = 0 To 13
        SQL = SQL & split_S(i) & " = '" & campos(i + 34)
        If i < 13 Then
        SQL = SQL & "',"
        Else:
        SQL = SQL & "' "
        End If
Next
n = Front1.ListBox1.ListIndex
If n < 0 Then
    Exit Sub
End If
SQL = SQL & "Where id = " & x_spl(n) & ""
BD.Open CS
BD.Execute SQL
BD.Close

Front1.MultiPage1.Value = 0
MsgBox "Alteração efetuada com sucesso!", vbInformation, "Informação"





End Sub
Sub update_cad()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, i As Integer
Dim split_S() As String, base As String


arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

base = "Nome,CPF,RG,Endereço,Bairro,Cidade,Estado," _
& "País,CEP,Complemento,Perfil,DataNasc,Telefone1,Telefone2,Email,Obs1," _
& "Nome2,CPF2,RG2,Endereço2,Bairro2,Cidade2,[Estado 2]," _
& "País2,CEP2,Complemento2,Perfil2,DataNasc2,Telefone12,Telefone22,Email2,Obs2,maioridade"


split_S() = Split(base, ",")



SQL = "UPDATE dadospessoais SET "
For i = 0 To 32
        SQL = SQL & split_S(i) & " = '" & campos(i + 1)
        If i < 32 Then
        SQL = SQL & "',"
        Else:
        SQL = SQL & "' "
        End If
Next
SQL = SQL & "Where nome = " & "'" & id & "'"

BD.Open CS
BD.Execute SQL



BD.Close

MsgBox "Alteração efetuada com sucesso!", vbInformation, "Informação"


End Sub

Sub clear_cadastro()
Dim i As Integer

For i = 1 To 32
    campos(i) = ""
Next

g = False
End Sub
Sub pesquisar()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, x As String, name As String, z As String
Dim c As Integer, n As Integer

arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

x = "%" & Front1.TextBox27 & "%"

SQL = "Select * from dadospessoais WHERE Nome LIKE "
SQL = SQL & "'" & x & "'"
SQL = SQL & " ORDER by nome ASC"

BD.Open CS
RS.Open SQL, BD

Front1.list123.Clear
x = ""
Do Until RS.EOF
    Front1.list123.AddItem RS.Fields(0) '& "     " & RS.Fields(1) & "     " & RS.Fields(10) & "     " & RS.Fields(14) & "     " & RS.Fields(15)
    x = x & RS.Fields(0) & ","
    RS.MoveNext
Loop

index1() = Split(x, ",")


RS.Close
BD.Close

End Sub

Sub dia_hoje()
Dim diasemana As String, x As Integer

x = Plan1Cb.Cells(3, 9).Value
diasemana = Plan1Cb.Cells(x, 10).Value
Plan1Cb.Range("i4").Value = diasemana

End Sub
Sub historic()

His_diario.ListBox1.Clear
Dim cntr As Integer, i As Integer, coluna As Integer

cntr = Application.WorksheetFunction.CountA(Range("O:O"))
For i = 2 To cntr
    His_diario.ListBox1.AddItem Plan1Cb.Cells.Range("Q2") & "                 " & Plan1Cb.Cells(i, 15) & "                 " & Plan1Cb.Cells(i, 16)
Next i
His_diario.Show
End Sub
Sub hist_e()
If Plan1Cb.Cells.Range("Q2") = Plan1Cb.Cells.Range("Q3") Then
    Exit Sub
Else:
    Dim cntr As Integer, i As Integer, j As Integer
    cntr = Application.WorksheetFunction.CountA(Range("O:O"))
    For i = 2 To cntr
        Plan1Cb.Cells(i, 15) = ""
        Plan1Cb.Cells(i, 16) = ""
    Next
    Plan1Cb.Cells.Range("Q4") = 2
    Plan1Cb.Cells.Range("Q3") = Plan1Cb.Cells.Range("Q2")
End If
End Sub
Sub atual_status()

Dim BD As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim arq As String, CS As String, SQL As String, n As Integer, x_2 As Integer

arq = ThisWorkbook.Path & "\" & "BD.accdb;"

CS = "Provider = Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & arq _
& "jet OLEDB:Database Password= 123456;"

n = Front1.listboxtodos.ListIndex
If n < 0 Then
    Exit Sub
End If
BD.Open CS

SQL = "SELECT * from especialidade"
SQL = SQL + " WHERE nome = '" & Front1.cbname1 & "'"
SQL = SQL + " AND especialidade = '" & Front1.ComboBox5 & "'"
SQL = SQL + " AND Status <> 'ABANDONO' AND Status <> 'ALTA' AND Status <> 'TRANSFERENCIA SAIDA'"

RS.Open SQL, BD

If RS.EOF = False Then
    x_2 = RS.Fields(15)
Else:
    Exit Sub
End If

RS.Close

SQL = "UPDATE especialidade SET "
If Front1.ComboBox10 = "ABANDONO" Or Front1.ComboBox10 = "ALTA" Or Front1.ComboBox10 = "TRANSFERENCIA SAIDA" Then
    SQL = SQL + "[Data Encerramento] = '" & Front1.datatxt & "',"
Else:
    Exit Sub
End If

SQL = SQL + " Status = '" & Front1.ComboBox10 & "'"
SQL = SQL + " WHERE id = " & x_2 & ""



BD.Execute SQL
BD.Close

End Sub
