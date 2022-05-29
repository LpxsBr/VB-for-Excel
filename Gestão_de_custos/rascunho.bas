Public Sub LBCMV()
a = ActiveCell.Address
b = Replace(a, "$", "")
'Debug.Print b

Dim COMP, EIM, EFM, CMV As Integer
Range(b).Offset(1, 0).Value = "RV"
Range(b).Offset(2, 0).Value = "(-)CMV"

Range(b).Offset(3, 0).Value = "EIM"
EIM = Range(b).Offset(3, 1).Value
Range(b).Offset(4, 0).Value = "COMP"
COMP = Range(b).Offset(4, 1).Value
Range(b).Offset(5, 0).Value = "EFM"
EFM = Range(b).Offset(5, 1).Value

CMV = (EIM + COMP) - EFM
Range(b).Offset(2, 1).Value = CMV
Range(b).Offset(6, 0).Value = "(=)LB"
Range(b).Offset(7, 0).Value = "(-)Dop"
Range(b).Offset(8, 0).Value = "(=)LL"

Debug.Print CMV
Debug.Print EIM
Debug.Print COMP
Debug.Print EFM

End Sub
