Attribute VB_Name = "Módulo1"
Sub finder()

name = InputBox("Digite abaixo", "Finder Box", "", 100, 100)

With ActiveSheet.Range("A:Z")

    Set tfinder = .Find(name, LookIn:=xlValues)

        Debug.Print "name = "; tfinder
        Debug.Print "location = "; tfinder.Address
        
End With

End Sub
