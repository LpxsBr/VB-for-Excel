Sub ForceRefresh()
'By: LpxsBr

  'calculate anything
Calculate

  'calculate all data of an active sheet
ActiveSheet.Calculate

  'activate automatic calculation
ActiveSheet.Application.Calculation = xlAutomatic
  
  'just a mensage
MsgBox ("Atualizado")

End Sub


