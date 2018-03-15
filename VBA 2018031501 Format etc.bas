Attribute VB_Name = "Ä£¿é1"
Sub FontSet031501()
'Format Prereduce
    'Font
    With Range("A1").CurrentRegion.Font
        .Name = "Î¢ÈíÑÅºÚ"
        .Size = 10
        .Color = RGB(0, 0, 0)
    End With
    'Borders
    With Range("A1").CurrentRegion.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlThin
    End With
    
    'First Row Interior Color Change into Grey
    Range("A1").CurrentRegion.Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Interior.Color = RGB(221, 221, 221)
    
    'Freeze the first row
    Range("A1").Select
    With ActiveWindow
      .SplitColumn = 0
      .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub
