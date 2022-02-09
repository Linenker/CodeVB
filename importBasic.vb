Sub ImportDB()

Application.ScreenUpdating = False

Dim vAltData               As Date
Dim vAltFicha              As String
Dim vAltCancel             As String
Dim vAltTransf             As String
Dim vAltCongel             As String
Dim vAltGuiche             As String
Dim vAltChegada            As Date
Dim vAltInicio             As Date
Dim vAltTME                As Date
Dim vAltFim                As Date
Dim vAltTMA                As Date
Dim vAltMatric             As String
Dim vAltUnidade            As String
Dim vAltUser               As String
Dim vAltFila               As String
Dim vAltPN                 As String
Dim vAltFone               As String
Dim vAltTitular            As String
Dim Base                   As Integer

Dim SQL             As String
Dim cn              As New ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim ln              As Integer
Dim col             As Integer
Dim cl              As Integer
Dim validf          As String
Dim W               As Worksheet
Dim resp            As VbMsgBoxResult

    Set W = Sheets("Base")
    W.Select
    vAltData = W.Range("A2")
    Base = 1
    Call validData(validf, Base)
    
        If validf = "No" Then
            MsgBox "Execução interrompda!"
            Exit Sub
        End If

Set cn = New ADODB.Connection
cn.Open conexaodb
Set RS = New ADODB.Recordset

lin = 2
col = 1
     
     Do Until W.Cells(lin, col) = ""
    
        vAltData = W.Cells(lin, col)
        vAltFicha = W.Cells(lin, col + 1)
        vAltCancel = W.Cells(lin, col + 2)
        vAltTransf = W.Cells(lin, col + 3)
        vAltCongel = W.Cells(lin, col + 4)
        vAltGuiche = W.Cells(lin, col + 5)
        vAltChegada = W.Cells(lin, col + 6)
        vAltInicio = W.Cells(lin, col + 7)
        vAltTME = W.Cells(lin, col + 8)
        vAltFim = W.Cells(lin, col + 9)
        vAltTMA = W.Cells(lin, col + 10)
        vAltMatric = W.Cells(lin, col + 11)
        vAltUnidade = W.Cells(lin, col + 12)
        vAltUser = W.Cells(lin, col + 13)
        vAltFila = Plan1.Cells(lin, col + 14)
        vAltPN = Plan1.Cells(lin, col + 15)
        vAltFone = Plan1.Cells(lin, col + 16)
        vAltTitular = Plan1.Cells(lin, col + 17)
                 
        SQL = "Insert into Banco_Dados_2022 " & _
              "(data, ficha, cancel, transf, congel, guiche, chegaCli, inicio, " & _
              "temp_esp, fim_atend, Temp_atend, matricula, unidade, usuario, fila, PN, Fone, Titular) "
        SQL = SQL & " values "
        SQL = SQL & "('" & vAltData & "'  , '" & vAltFicha & "' , '" & vAltCancel & "' , "
        SQL = SQL & "'" & vAltTransf & "' , '" & vAltCongel & "', '" & vAltGuiche & "' , "
        SQL = SQL & "'" & vAltChegada & "', '" & vAltInicio & "', '" & vAltTME & "'    ,"
        SQL = SQL & "'" & vAltFim & "'    , '" & vAltTMA & "'   , '" & vAltMatric & "' ,"
        SQL = SQL & "'" & vAltUnidade & "', '" & vAltUser & "'  , '" & vAltFila & "' ,"
        SQL = SQL & "'" & vAltPN & "', '" & vAltFone & "'  , '" & vAltTitular & "'   )"
        
        On Error GoTo analize
        RS.Open SQL, cn
        
     lin = lin + 1
     Loop

    cn.Close
    Application.ScreenUpdating = False
    Exit Sub
   
analize:
        W.Range("A" & lin).Select
        resp = MsgBox("Erro no processo de importação na linha  " & lin & "!, deseja continuar?", vbYesNo + vbQuestion, "Analise de Erro!")
        
            If resp = vbYes Then
                    cn.Close
                    Exit Sub
            Else
                MsgBox "A base será apagada no banco de dados!", vbExclamation, "Analise de Erro!"
                
                Call ExcluirDados(Base)
            End If
End Sub
