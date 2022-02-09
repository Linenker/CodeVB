'*- Exportar dados do banco de dados
Dim vUltCelSel  As Range
Sub ImportacaoDeDados()
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
    Set W = Sheets("DB")
    W.Select
    W.Range("A6").Select
    col = 1
    
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
            With W.Cells(5, col)
            
            .Value = FD.Name
            .Font.Bold = True
            .Interior.Color = RGB(37, 219, 119)
            End With
        col = col + 1
    Next FD

    '*- Exportar dados do Banco de dados
        W.Cells(6, 1).CopyFromRecordset RS
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

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub