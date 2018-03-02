Attribute VB_Name = "模块4"
Sub ShowAllFiles030203()
    '读取当前文件夹下所有文件
    Dim myPath As String
    Dim myFileName As String
    Dim i As Long
    'myPath = ThisWorkbook.Path & "\"
    myPath = InputBox("Please input file path you want to read.", "Get Path") & "\"
    '指定文件夹,目前为读取输入值
    myFileName = Dir(myPath, 0)
    i = 0
    Do While Len(myFileName) > 0
        Cells(i + 1, 1) = myPath & myFileName
        myFileName = Dir()
        i = i + 1
    Loop
End Sub

