'serao usados para construção de suplementos
'Arquivo deve ser salvo como .XLAM (suplemento de excel) e adicionado a pasta C:\Users\usuario\AppData\Roaming\Microsoft\AddIns

'Para usar
'Abra o excel
'vá em opçoes
'lá em baixo, Suplementos
'lá em baixo, gerir: suplementos do Excel
'clica em "ir"
'vai abrir um menu, marca o nome do suplemento e dá ok
'pronto pra usar

'senão achar, clica em procurar e vai na sua pasta de interesse

'QUANDO FOR USAR E QUISER VER UMA DESCRIÇÃO DE COMO USAR A FORMULA
'DIGITE A FORMULA ATÉ PARENTESE (OU DÊ UM TAB) E TECLE, CTRL+SHIFT+A

'EX: =CPP( CTRL+SHIFT+A

Public Function LLCustAbs(ByVal receita_venda As Double, ByVal custo_produto_vendido As Double, ByVal despesa_operacional_periodo As Double)
    
    LLCustAbs = receita_venda - custo_produto_vendido - despesa_operacional_periodo

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------

Sub lsDefinirCategoria()
    Application.MacroOptions Macro:="LLCustAbs", Category:=1
End Sub
