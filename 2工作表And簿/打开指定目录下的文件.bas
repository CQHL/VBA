Attribute VB_Name = "ģ��1"


Sub ��ָ��Ŀ¼�µ��ļ�()
Dim a$, n!, wbs As Workbook
a = Dir("C:\*.xls")
Workbooks.Open "C:\" & a
Do
    a = Dir
    If a <> "" Then
        Workbooks.Open "C:\" & a
    Else
        Exit Sub
    End If
Loop
End Sub
