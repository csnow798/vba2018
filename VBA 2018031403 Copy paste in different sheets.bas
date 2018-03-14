Attribute VB_Name = "Ä£¿é8"
Sub Copypaste031403()
'Copy A1:A10 in Sheet11 to B1:B10 in Sheet8
    Worksheets("Sheet11").Range("A1:A10").Copy Destination:=Worksheets("Sheet8").Range("B1:B10")
End Sub
