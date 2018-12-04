Attribute VB_Name = "Ä£¿é1"
Sub ¾Å¾Å³Ë·¨±í()
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
