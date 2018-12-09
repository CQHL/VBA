Attribute VB_Name = "模块1"


Sub 打开指定目录下的文件()
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
