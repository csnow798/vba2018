Attribute VB_Name = "ģ��5"
Sub IsOpen031507()
    Dim i As Integer
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "CloudBU" Then
            '�ƶ�����������֮ǰ
            Worksheets(i).Move before:=Worksheets(1)
            MsgBox "Sheet Found"
            Exit Sub
        End If
    Next
    Worksheets.Add before:=Worksheets(1)
    Worksheets(1).Name = "CloudBU"
    MsgBox "404 Not Found"
End Sub
