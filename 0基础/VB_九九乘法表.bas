Attribute VB_Name = "ģ��1"
Sub �žų˷���()
Dim a!, b!
For a = 1 To 9
    For b = 1 To 9
        If b > a Then
            Sheet1.Cells(a, b) = ""
        Else
            Sheet1.Cells(a, b) = a & "X" & b & "=" & a * b
        End If
       Next
    Next
End Sub
