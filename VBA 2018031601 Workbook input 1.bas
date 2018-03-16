Attribute VB_Name = "模块6"
Sub WbInput()
    Dim wb As String, xrow As Integer, arr
    wb = ThisWorkbook.Path & "\员工花名册.xlsx"
    Workbooks.Open (wb)
    With ActiveWorkbook.Worksheets(1)
    xrow = .Range("A1").CurrentRegion.Rows.Count + 1
    arr = Array(xrow - 1, "ZhangJiao", "Female", #7/8/1987#, #9/1/2010#, "2010 New Hired")
    .Cells(xrow, 1).Resize(1, 6) = arr
    End With
    ActiveWorkbook.Close savechanges:=True
End Sub
