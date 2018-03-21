Attribute VB_Name = "Ä£¿é3"
Sub ColorIndex032104()
    'Cells(1, 1) = "#"
    'Cells(1, 2) = "Color"
    'Range("A1:B1").Interior.ColorIndex = 15
    Range(Cells(1, 1), Cells(2, 56)).Interior.ColorIndex = xlNone
    Dim i As Integer
    i = 1
    For i = 1 To 56
        Cells(i, 1).Value = i
        Cells(i, 2).Interior.ColorIndex = i
    Next
End Sub

