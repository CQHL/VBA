Sub 多表合并()
Dim i%,rs%, rss% st As Worksheetm ,zst As Worksheet
Set zst = Sheets("1季度")
For i = 1 To 3
	Set st = Sheets(i & "月")
	rs = st.usedRange.Rows.Count
	rss = zst.usedRange.Rows.Count + 1
	st.Range("a2:b" & rs).Copy Cells(rss, 1)
	Cells(rss,3).Resize(rs-1) = i & "月"
Next
End Sub