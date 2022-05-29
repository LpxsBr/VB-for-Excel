'suplementos simples

'Adicione o suplemento nessa pasta C:\Users\usuario\AppData\Roaming\Microsoft\AddIns

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

Public Function CPP(ByVal CD As Double, ByVal CIF As Double)
    CPP = CD + CIF
End Function

Public Function CPP2(ByVal CF As Double, ByVal CV As Double)
    CPP = CV + CF
End Function

Public Function CPP3(ByVal Mat_Direto As Double, ByVal Mao_Obra_Direta As Double, ByVal CIF As Double)
    CPP = Mat_Direto + Mao_Obra_Direta + CIF
End Function

Function CD(ByVal Mat_Direto As Double, ByVal Mao_Obra_Direta As Double)
    CD = Mat_Direto + Mao_Obra_Direta
End Function

Public Function CPPUnit(ByVal CPP As Double, ByVal Quant_Produzida As Double)
    CPPUnit = CPP / Quant_Produzida
End Function

Public Function ConsMP(ByVal EIMP As Double, ByVal COMP As Double, ByVal EFMP As Double)
    ConsMP = EIMP + COMP - EFMP
End Function
Public Function EFMP(ByVal EIMP As Double, ByVal COMP As Double, ByVal ConsMP As Double)
    EFMP = EIMP + COMP - ConsMP
End Function

Public Function CPA(ByVal EIPE As Double, ByVal CPP As Double, ByVal EFPE As Double)
    CPA = EIPE + CPP - EFPE
End Function

Public Function EIPE(ByVal CPA As Double, ByVal EFPE As Double, ByVal CPP As Double)
    EIPE = CPA + EFPE - CPP
End Function
Public Function EFPE(ByVal EIPE As Double, ByVal CPP As Double, ByVal CPA As Double)
    EFPE = EIPE + CPP - CPA
End Function

Public Function CPV(ByVal EIPE As Double, ByVal CPA As Double, ByVal EFPA As Double)
    CPV = EIPA + CPA - EFPA
End Function

Public Function LB(ByVal Receita_de_venda As Double, ByVal CPV As Double)
    LB = Receita_de_venda - CPV
End Function

Public Function LL_NORMAL(ByVal LB As Double, ByVal Despesa_OP As Double)
    LL_NORMAL = LB - Despesa_OP
End Function

Public Function CMV(ByVal EstoqueI_Mercadoria As Double, ByVal Compra_mercadoria As Double, EstoqueF_Mercadoria As Double)
    CMV = EstoqueI_Mercadoria + Compra_mercadoria - EstoqueF_Mercadoria
End Function

Function COLADASFORMULADECUSTOS() As String

MsgBox ("CPP = CD + CIF = MD + MOD + CIF " & Chr(13) & Chr(13) & " CPP = CF + CV " & Chr(13) & Chr(13) & " CD = MD + MOD " & Chr(13) & Chr(13) & "  CPP = CPPu . Qp " & Chr(13) & Chr(13) & "  CTr = MOD + CIF " & Chr(13) & Chr(13) & "MD = CONS(MP) = EIMP + COMP - EFMP" & Chr(13) & Chr(13) & "CPA = EIPE + CPP - EFPE" & Chr(13) & Chr(13) & "CPV = EIPA + CPA - EFPA" & Chr(13) & Chr(13) & "LB = RV - CPV " & Chr(13) & Chr(13) & " EFMP = EIMP + COMP - CONS" & Chr(13) & Chr(13) & "CONS(MP) = MD = EIMP + COMP - EFMP" & Chr(13) & Chr(13) & "CPP = MD + MOD + CIF" & Chr(13) & Chr(13) & "EIPE CPP - CPA = EFPE" & Chr(13) & Chr(13) & "CPA = EIPE + CPP - EFPE" & Chr(13) & Chr(13) & "EIPA CPA - CPV = EFPA" & Chr(13) & Chr(13) & "CPV = EIPA + CPA - EFPARV -CPV = LB" & Chr(13) & Chr(13) & "LB -DOp = LL")
MsgBox ("RV" & Chr(13) & Chr(13) & "(-)CMV" & Chr(13) & Chr(13) & "EIM" & Chr(13) & Chr(13) & "COMP" & Chr(13) & Chr(13) & "EFM" & Chr(13) & Chr(13) & "(=)LB" & Chr(13) & Chr(13) & "(-)Dop" & Chr(13) & Chr(13) & "(=)LL")
MsgBox ("VALOR S/ ICMS E S/ IPI" & Chr(13) & Chr(13) & "(+) ICMS" & Chr(13) & Chr(13) & "(=) VM" & Chr(13) & Chr(13) & "(+) IPI" & Chr(13) & Chr(13) & "(=) VNF" & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "V S/ ICMS + Aliquota.VM = VM" & Chr(13) & Chr(13) & "V S/ ICMS = VM . (1-Aliquota)")

End Function



