Attribute VB_Name = "ģ��9"
Sub PasteValueOnly031401()
'
' PasteValueOnly031401 ��
'
    Range("A1:D10").Copy
    Range("F1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


