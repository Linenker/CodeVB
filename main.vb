Sub MainAjuste()

    Call Ajuste_Base
    Call Ajuste_Base_SM
    Call Ajuste_Base_LL
    Call Ajuste_Pesquisa_Satisfacao


End Sub
Sub MainImport()
    
    Call ImportDB
    Call ImportSM
    Call Import_Login_Logout
    Call Import_Base_Pesquisa

End Sub
Sub MainLimpa()
    Call LimpaBase
    Call LimpaBaseSM
    Call LimpaBaseLL
    Call LimparBasePesquisa
End Sub
Sub CleanEspecializado()
        Plan9.Range("E5:G7").Select
        Selection.ClearContents
End Sub
Sub CleanEspecializadoAt()
        Plan9.Range("B10:M26").Select
        Selection.ClearContents
End Sub