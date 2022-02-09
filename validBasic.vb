Function validData(validf, Base)
Dim SQL                 As String
Dim cn                  As New ADODB.Connection
Dim RS                  As New ADODB.Recordset
Dim dataValid           As String
Dim resp                As VbMsgBoxResult


    'Conexão banco de dados
        Set cn = New ADODB.Connection
        cn.Open conexaodb
        Set RS = New ADODB.Recordset


        Select Case Base
Case 1
        dataValid = Format(Plan1.Range("A2"), "mm/dd/yyyy")
        SQL = "SELECT tb.data, * FROM Banco_Dados_2022 AS tb WHERE (((tb.data)=#" & dataValid & "#));"
        
Case 2
        dataValid = Format(Plan2.Range("A2"), "mm/dd/yyyy")
        SQL = "Select tb.data, * FROM DB_SenhaMesa_2022 AS tb WHERE (((tb.data)=#" & dataValid & "#));"
    
Case 3
        dataValid = Format(Plan5.Range("A2"), "mm/dd/yyyy")
        SQL = "Select tb.data, * FROM LoginLogout AS tb WHERE (((tb.data)=#" & dataValid & "#));"
Case 4
        Call ExcluirDados(Base)
        Exit Function

Case 5
        dataValid = Format(Plan7.Range("A2"), "mm/dd/yyyy")
        SQL = "Select tb.data, * FROM Pesquisa_Satisfacao_2022 AS tb WHERE (((tb.data)=#" & dataValid & "#));"
        
Case 6
        dataValid = Format(PSAJUSTE.Range("B2"), "mm/dd/yyyy")
        SQL = "Select tb.data, * FROM PesquisaSatisfacao_2022 AS tb WHERE (((tb.data)=#" & dataValid & "#));"
        
        End Select
        
    
        RS.Open SQL, cn
        If RS.EOF = True Then
                RS.Close
                cn.Close
            Exit Function
        Else
             resp = MsgBox("Data já consta no sistema, deseja substituir?", vbYesNo + vbExclamation, "Validação de Data!")
             If resp = vbNo Then
                    RS.Close
                    cn.Close
                    validf = "No"
                    Exit Function
             Else
                Call ExcluirDados(Base)
            End If
        End If
        
End Function
Function ExcluirDados(Base)
    Dim cn          As New ADODB.Connection
    Dim RS          As New ADODB.Recordset
    Dim FD          As ADODB.Field
    Dim SQL         As String
    Dim dataDelet   As String
    
    Set cn = New ADODB.Connection
    cn.Open conexaodb
    
    
        Select Case Base
Case 1
            dataDelet = Format(Plan1.Range("A2"), "mm/dd/yyyy")
            SQL = "DELETE tb.data, tb.* FROM Banco_Dados_2022 AS tb WHERE (((tb.data)=#" & dataDelet & "#));"
Case 2
            dataDelet = Format(Plan2.Range("A2"), "mm/dd/yyyy")
            SQL = "DELETE tb.data, tb.* FROM DB_SenhaMesa_2022 AS tb WHERE (((tb.data)=#" & dataDelet & "#));"
Case 3
            dataDelet = Format(Plan5.Range("A2"), "mm/dd/yyyy")
            SQL = "DELETE tb.data, tb.* FROM LoginLogout AS tb WHERE (((tb.data)=#" & dataDelet & "#));"
Case 4
            Exit Function
Case 5
            dataDelet = Format(Plan7.Range("A2"), "mm/dd/yyyy")
            SQL = "DELETE tb.data, tb.* FROM Pesquisa_Satisfacao_2022 AS tb WHERE (((tb.data)=#" & dataDelet & "#));"
Case 6
            dataDelet = Format(PSAJUSTE.Range("A2"), "mm/dd/yyyy")
            SQL = "DELETE tb.data, tb.* FROM PesquisaSatisfacao_2022 AS tb WHERE (((tb.data)=#" & dataDelet & "#));"
        
        End Select
    
    On Error GoTo exit_point
            cn.Execute SQL
            cn.Close
            Exit Function
    
exit_point:
    On Error Resume Next
    cn.Close
    MsgBox "Data não localizada na celula A2!"
    
End Function
