Sub insereFormulaComplete()
'1
umalocalizacao = (ActiveSheet.Range("A1").End(xlDown).Address)
a = Replace(umalocalizacao, "$A$", "")
numerodeumalocalizacao = CInt(a)

'2.1
outralocalizacao = (ActiveSheet.Range("J2").End(xlDown).Address)
f = Replace(outralocalizacao, "$J$", "")
numerodeoutralocalizacao = CInt(f)


Range(Cells(numerodeoutralocalizacao, 10), Cells(numerodeoutralocalizacao, 19)).Select
Selection.AutoFill Destination:=Range(Cells(numerodeoutralocalizacao, 10), Cells(numerodeumalocalizacao, 19))
Cells(numerodeoutralocalizacao, 1).Select
    
End Sub

