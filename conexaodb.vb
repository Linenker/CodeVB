Function conexaodb()

    Dim Arq                         As String
    
    Arq = ActiveWorkbook.Path & "\2022_Banco_Geral.accdb"
    conexaodb = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Arq & ";Persist Security Info=False"
    
End Function