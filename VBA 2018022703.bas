Attribute VB_Name = "Ä£¿é5"
Function countcolor(arr As Range)
    Dim rng As Range
    For Each rng In arr
        If rng.Interior.Color = RGB(255, 255, 0) Then
            countcolor = countcolor + 1
        End If
    Next rng
End Function
