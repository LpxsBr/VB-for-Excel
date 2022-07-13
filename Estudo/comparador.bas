Sub compara()

Dim ultimalinha As LongLong
Dim x As LongLong
Dim y As Integer

a = ActiveSheet.Range("A1").End(xlDown).Address
b = Replace(a, "$A", "")

'variavel 1
ultimalinha = (CInt(b) + 0)

'filtro da primeira aba

Worksheets("relatorio de vendas").Range("A:Z").AutoFilter Field:=21, Criterial:= _
        "Atendido"
        
x = 1
' posição da coluna de pedidos
y = 22

a = Worksheets("relatorio de vendas").Cells(x, y).Value
b = Worksheets("relatorio de vendas (2)").Cells(x, y).Value

'While x < ultimalinha
'    If a <> b Then    
    
'    End If
'Wend

Debug.Print a
Debug.Print b
Debug.Print ultimalinha

End Sub
