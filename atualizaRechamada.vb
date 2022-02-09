'*- Exportar dados do banco de dados
Dim vUltCelSel  As Range
Sub Rechamada()
Set vUltCelSel = ActiveCell

Application.EnableEvents = False
Application.ScreenUpdating = False

'*- Atribuir Variaveis
'Acionar o pacote Microsoft ActiveX Data Objects 2.8 library // DLL raiz do da Microsoft para programação em VBA
Dim SQL         As String
Dim cn          As New ADODB.Connection
Dim RS          As New ADODB.Recordset
Dim FD          As ADODB.Field
Dim col         As Integer
Dim W           As Worksheet

'*- Manipulção da planilha
    Set W = Sheets("Rechamada")
    W.Select
    W.Range("A6").Select
    col = 1
    
    W.Range("A2:D1048576").ClearContents
    
'*- Criar a conexão
Set cn = New ADODB.Connection

'*- Abrir a conexão
cn.Open conexaodb '(Function criada para com a rota do banco de dados)

'*- Recordset
Set RS = New ADODB.Recordset

'*- Consulta Principal
SQL = RetornaSQL(1)
'*- Abrir a Consulta
RS.Open SQL, cn
'*- Verificar a se tem dados na consulta
If RS.EOF = False Then
    
    '*- Adicionar o nome na colunas
        For Each FD In RS.Fields
            With W.Cells(1, col)
            
            .Value = FD.Name
            .Font.Bold = True
            .Interior.Color = RGB(37, 219, 119)
            End With
        col = col + 1
    Next FD

    '*- Exportar dados do Banco de dados
        W.Cells(2, 1).CopyFromRecordset RS
        Application.StatusBar = "Consulta concluida com sucesso..."
Else
    MsgBox "Não há dados para serem trazidos..."
End If

'*- Fechar o Recordset
RS.Close

'*- Fechar o Banco de dados
cn.Close

    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Call FormulaRechamada
Call ImportRechamada

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub
'*- Criar uma function de SQL
Function RetornaSQL(vQuery As String)

Dim MesBase                         As String
Dim W                               As Worksheet


Select Case vQuery

Case 1
    SQL = "SELECT Banco_Dados_2022.data, Banco_Dados_2022.unidade, " & _
    "Banco_Dados_2022.Fone, Count(Banco_Dados_2022.Fone) AS Qnt " & _
    "FROM Banco_Dados_2022 GROUP BY Banco_Dados_2022.data, " & _
    "Banco_Dados_2022.unidade, Banco_Dados_2022.Fone;"

Case 2

SQL = ""
    
End Select

RetornaSQL = SQL

End Function
Sub FormulaRechamada()
Application.Calculation = xlCalculationManual
lin = 2

Do Until Plan9.Range("A" & lin) = ""
Plan9.Select
Plan9.Range("E" & lin).Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",IFERROR(RC[-2]*1=0,0)),""Sem registro"",IF(RC[-1]>1,""Rechamada"",""Contato Único""))"
lin = lin + 1
Loop
Application.Calculation = xlCalculationAutomatic
End Sub
Sub ImportRechamada()
Application.ScreenUpdating = False


Dim at_data              As Date
Dim at_unidade               As String
Dim at_Fone              As String
Dim at_Qnt               As String
Dim at_Rechamada         As String


Dim SQL             As String
Dim cn              As New ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim ln              As Integer
Dim col             As Integer
Dim cl              As Integer
Dim validf          As String
Dim W               As Worksheet
Dim resp            As VbMsgBoxResult

    Set W = Sheets("Rechamada")
    W.Select

Set cn = New ADODB.Connection
cn.Open conexaodb
Set RS = New ADODB.Recordset
            
    SQL = "DELETE Rechamada_2022.* FROM Rechamada_2022;"
    cn.Execute SQL

lin = 2
col = 1
     
     Do Until W.Cells(lin, col) = ""

     at_data = Format(W.Range("A" & lin), "dd/mm/yyyy")
     at_unidade = W.Range("B" & lin)
     at_Fone = W.Range("C" & lin)
     at_Qnt = W.Range("D" & lin)
     at_Rechamada = W.Range("E" & lin)
                    
        SQL = "Insert into Rechamada_2022 " & _
              "(Data, Unidade, Fone, Qnt, Rechamada) "
        SQL = SQL & " values "
        SQL = SQL & "( '" & at_data & "' , '" & at_unidade & "' , '" & at_Fone & "' , '" & at_Qnt & "' , '" & at_Rechamada & "' )"
        'Range("F1") = SQL
        
        RS.Open SQL, cn
        
     lin = lin + 1
     Loop

    cn.Close
    Application.ScreenUpdating = False
End Sub
