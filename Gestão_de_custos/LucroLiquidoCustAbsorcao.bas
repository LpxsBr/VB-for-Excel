'serao usados para construção de suplementos

Public Function LLCustAbs(ByVal receita_venda As Double, ByVal custo_produto_vendido As Double, ByVal despesa_operacional_periodo As Double)
    
    LLCustAbs = receita_venda - custo_produto_vendido - despesa_operacional_periodo

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------

Sub lsDefinirCategoria()
    Application.MacroOptions Macro:="LLCustAbs", Category:=1
End Sub
