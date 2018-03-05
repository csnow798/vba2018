VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub choosevaluesA1()
    'Select range if it value equals to "A1"
    Dim myrange As Range, n As Range
    Set myrange = Range("A1")
    For Each n In Range("A1:D10")
        If n.Value = Range("A1").Value Then
            Set myrange = Union(myrange, n)
        End If
    Next
    myrange.Select
End Sub
