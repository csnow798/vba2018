Attribute VB_Name = "ģ��4"
Sub ShowAllFiles030203()
    '��ȡ��ǰ�ļ����������ļ�
    Dim myPath As String
    Dim myFileName As String
    Dim i As Long
    'myPath = ThisWorkbook.Path & "\"
    myPath = InputBox("Please input file path you want to read.", "Get Path") & "\"
    'ָ���ļ���,ĿǰΪ��ȡ����ֵ
    myFileName = Dir(myPath, 0)
    i = 0
    Do While Len(myFileName) > 0
        Cells(i + 1, 1) = myPath & myFileName
        myFileName = Dir()
        i = i + 1
    Loop
End Sub

