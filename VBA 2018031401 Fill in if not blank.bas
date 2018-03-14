Attribute VB_Name = "Ä£¿é5"
Sub Rngend031401()
    'Fill in data into the first cell which is not blank on column A
    If ActiveSheet.Range("A65535").End(xlUp) = "" Then
        ActiveSheet.Range("A65535").End(xlUp).Value = "Fill in when Blank"
    Else
        ActiveSheet.Range("A65535").End(xlUp).Offset(1, 0).Value = "Not Blank, Fill into next row"
    End If
End Sub
