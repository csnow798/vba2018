Attribute VB_Name = "Ä£¿é9"
Sub PasteValueOnly031401()
'
' PasteValueOnly031401 ºê
'
    Range("A1:D10").Copy
    Range("F1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


