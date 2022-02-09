'*- Criar uma function de conexão ao banco
Function conexaodb()

    Dim arq                         As String
    
    arq = ActiveWorkbook.Path & "\2022_Banco_Geral.accdb"
    conexaodb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & arq & ";Persist Security Info=False"
    
End Function
'*- Criar uma function de SQL
Function RetornaSQL(vQuery As String)
Dim MesBase                         As String
Dim W                               As Worksheet

    Set W = Sheets("Config")
    MesBase = W.Range("G2")


Select Case vQuery
Case 1
    SQL = "SELECT Format([data],'mmmm') AS Mes, Banco_Dados_2022.data, Banco_Dados_2022.cancel, " & _
    "Max([temp_esp]*86400) AS TME, Avg([Temp_atend]*86400) AS TMA, Regiao.Regiao, " & _
    "Banco_Dados_2022.unidade, Banco_Dados_2022.fila, Count(Banco_Dados_2022.fila) " & _
    "AS TotalAtend, Fila.NS, Fila.Maximo, Fila.AjustTME, Fila.AjustTMA " & _
    "FROM Fila INNER JOIN (Regiao INNER JOIN Banco_Dados_2022 ON Regiao.Cidade = Banco_Dados_2022.unidade) " & _
    "ON Fila.Nome_Fila = Banco_Dados_2022.fila " & _
    "GROUP BY Format([data],'mmmm'), Banco_Dados_2022.data, Banco_Dados_2022.cancel, " & _
    "Regiao.Regiao, Banco_Dados_2022.unidade, Banco_Dados_2022.fila, Fila.NS, " & _
    "Fila.Maximo, Fila.AjustTME, Fila.AjustTMA " & _
    "HAVING (((Format([data],'mmmm'))='" & MesBase & "'));"

        
Case 2
      SQL = "SELECT Format([data],'mmmm') AS MesBase, Banco_Dados_2022.data, Regiao.Regiao, " & _
        "Banco_Dados_2022.unidade, Banco_Dados_2022.fila, Count(Banco_Dados_2022.usuario) " & _
        "AS ContarDeusuario, Sum(IIf([Temp_esp]*86400>1620,0,1)) AS Ns27, Sum(IIf([Temp_esp]*86400>1800,0,1)) " & _
        "AS Ns30, Sum(IIf([Temp_esp]*86400>2400,0,1)) AS Ns40, Sum(IIf([Temp_esp]*86400>600,0,1)) " & _
        "AS Ns10, Sum(IIf([Temp_esp]*86400>1200,0,1)) AS Ns20 " & _
        "FROM (Regiao INNER JOIN Banco_Dados_2022 ON Regiao.Cidade = Banco_Dados_2022.unidade) " & _
        "INNER JOIN Fila ON Banco_Dados_2022.fila = Fila.Nome_Fila " & _
        "GROUP BY Banco_Dados_2022.data, Regiao.Regiao, Banco_Dados_2022.unidade, Banco_Dados_2022.fila, Banco_Dados_2022.cancel, Fila.NS " & _
        "HAVING (((Format([data],'mmmm'))='" & MesBase & "') AND ((Banco_Dados_2022.cancel)<>'S') AND ((Fila.NS)='S'));"

    Range("A1") = SQL
     
    
Case 3
    SQL = "SELECT Format([Data],'mmmm') AS MesBase, DB_SenhaMesa_2022.Data, Regiao.Regiao, " & _
        "DB_SenhaMesa_2022.Unidade, DB_SenhaMesa_2022.Serviço, DB_SenhaMesa_2022.TA, " & _
    "DB_SenhaMesa_2022.TMA, DB_SenhaMesa_2022.Qtd " & _
    "FROM DB_SenhaMesa_2022 INNER JOIN Regiao ON DB_SenhaMesa_2022.Unidade = Regiao.Cidade " & _
    "GROUP BY Format([Data],'mmmm'), DB_SenhaMesa_2022.Data, Regiao.Regiao, " & _
    "DB_SenhaMesa_2022.Unidade, DB_SenhaMesa_2022.Serviço, DB_SenhaMesa_2022.TA, " & _
    "DB_SenhaMesa_2022.TMA, DB_SenhaMesa_2022.Qtd " & _
    "HAVING (((Format([Data],'mmmm'))='" & MesBase & "'));"
   
    
Case 4
    MesBase = Format(W.Range("G3"), "MM")
    SQL = "SELECT Format([Data],'mm') AS MB, Pesquisa_Satisfacao_2022.Data, " & _
    "Count(Pesquisa_Satisfacao_2022.Unidade) AS Qnt, Pesquisa_Satisfacao_2022.Unidade, " & _
    "Regiao.Regiao, Avg(Pesquisa_Satisfacao_2022.GrauSatisfacao) AS MédiaDeGrauSatisfacao, " & _
    "Avg(Pesquisa_Satisfacao_2022.RecomendarEDP) AS MédiaDeRecomendarEDP, Pesquisa_Satisfacao_2022.Resolvido " & _
    "FROM Pesquisa_Satisfacao_2022 INNER JOIN Regiao ON Pesquisa_Satisfacao_2022.Unidade = Regiao.Cidade " & _
    "GROUP BY Format([Data],'mm'), Pesquisa_Satisfacao_2022.Data, " & _
    "Pesquisa_Satisfacao_2022.Unidade, Regiao.Regiao, Pesquisa_Satisfacao_2022.Resolvido " & _
    "HAVING (((Format([Data],'mm'))='" & MesBase & "'));"
    
    
    Case 5
    SQL = "SELECT Format([data],'mm_mmmm') AS mes, Banco_Dados_2022.Fone, " & _
    "Count(Banco_Dados_2022.Fone) AS Qnt FROM Banco_Dados_2022 " & _
    "GROUP BY Format([data],'mm_mmmm'), Banco_Dados_2022.Fone;"
    
End Select

RetornaSQL = SQL

End Function
Sub MainGeral()
Application.Calculation = xlCalculationManual
    
    Call CleanGeral
    Call ImportacaoDeDados
    Call ImportacaoNS
    Call ImportacaoSM
    Call MainFormat
    Call Dashboard
    Call Export_Pesquisa
    
Application.Calculation = xlCalculationAutomatic
End Sub
Sub CleanGeral()
        Plan5.Range("A6:CX1048576").ClearContents
        Plan7.Range("A6:CX1048576").ClearContents
        Plan8.Range("A6:CX1048576").ClearContents
        
End Sub