Attribute VB_Name = "Ä£¿é6"
Sub Rngend031402()
    'Fill in data into the first cell which is not blank on column A
    'Reconstruction with Uedrange
    If ActiveSheet.Range("A65535").End(xlUp) = "" Then
        ActiveSheet.UsedRange.Resize(1, 1).Value = "Fill in when Blank"
    Else
        ActiveSheet.UsedRange.End(xlDown).Offset(1, 0).Value = "Not Blank, Fill into next row"
    End If
End Sub

