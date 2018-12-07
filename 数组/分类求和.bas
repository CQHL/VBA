
Sub 计算不重复值列并统计（）
Dim arr1(1 To 10, 1 To 2)
Set endr = Cells(Rows.Count,"c").End(xlUp)
arr = Range([b2],endr)
For i = 1 To endr.Row - 1
	For j = 1 To UBound(arr1)
		x = arr(i,1):y = arr1(j,1)
		if arr(i,1) = arr1(j,1) Then
			arr1(j , 2) = arr( i , 2) + arr1(j + 2)
			GoTo 100
		End If
	Next j
		k = k + 1
		arr1(k, 1) = arr(i, 1)
		arr1(k, 2) = arr(i, 2)
100:
Next i
[e2].Resize(k, 2) = arr1
End Sub
