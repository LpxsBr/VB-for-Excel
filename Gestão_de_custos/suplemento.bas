Public Function CPP(ByVal custo_direto As Double, ByVal Cust_Indireto_Fabricacao As Double)

    CPP = custo_direto + Cust_Indireto_Fabricacao
    
End Function
Public Function CPP2(ByVal Mat_direto As Double, ByVal Mao_Obra_Direta As Double, ByVal Cust_Indireto_Fabricacao As Double)

    CPP = Mat_direto + Mao_Obra_Direta + Cust_Indireto_Fabricacao
    
End Function

Public Function LB(ByVal RV As Double, ByVal Cust_Prod_Vendido As Double)

    CPV = Cust_Prod_Vendido

    LB = RV - CPV

End Function

Public Function EFMP(ByVal EIMP As Double, ByVal Compras As Double, ByVal Consumo_MP As Double)

    EFMP = EIMP + Compras - Consumo_MP

End Function

Public Function CPA(ByVal EIPE As Double, ByVal CPP As Double, ByVal EFPE As Double)

    CPA = EIPE + CPP - EFPE

End Function
