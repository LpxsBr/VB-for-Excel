Sub numeradorLateral()

Dim ultimalinha As LongLong
Dim x As LongLong

a = (ActiveSheet.Range("A1").End(xlDown).Address)

b = Replace(a, "$A$", "")

'variavel 1
ultimalinha = (CInt(b))

'variavel de contagem começando da ultima numeracao preenchida
a = (ActiveSheet.Range("B1").End(xlDown).Address)
b = Replace(a, "$B$", "")
'variavel 1
x = (CInt(b) + 1)

'variavel de posição da coluna
y = 2
While x < ultimalinha
    x = x + 1
    ActiveSheet.Cells(x, y).Value = x
    Debug.Print x
Wend
End Sub
