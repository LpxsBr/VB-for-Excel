Sub insertCalcInPlan()

'firt reference
'defining the reference of the last name added in the worksheet
localRootOne = (ActiveSheet.Range("A1").End(xlDown).Address)
a = Replace(localRootOne, "$A$", "")

'convert the number in Integer for to use in Cells Method how reference
localConverted = CInt(a) 

'the same thing of firt reference
localTwoFinal = (ActiveSheet.Range("J2").End(xlDown).Address)
f = Replace(localTwoFinal, "$J$", "")

'convert the number in Integer for to use in Cells Method how reference
useLocal = CInt(f) 

'using range method to select the space of cells to do the operation of autofill in excel worksheet
Range(Cells(useLocal, 10), Cells(useLocal, 19)).Select

'destination of autofill
Selection.AutoFill Destination:=Range(Cells(useLocal, 10), Cells(localConverted, 19))
Cells(useLocal, 1).Select
    
End Sub

