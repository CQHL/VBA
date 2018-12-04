
Sub chr函数数字符循环()
Dim r%, c%
For i = 1 To 65535
    r = (i - 1) \ 10 + 1
    c = 2 * ((i - 1) Mod 10) + 1
    Cells(r, c) = i
    Cells(r, c + 1) = Chr(i)
Next
End Sub
