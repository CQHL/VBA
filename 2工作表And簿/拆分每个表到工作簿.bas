Attribute VB_Name = "ģ��1"


Sub ��ֵ�������()
Dim wk As Workbook, ss$, k%
Application.DisplayAlerts = False
For Each sht In ThisWorkbook.Sheets
    Set wk = Workbooks.Add
    k = k + 1
    Workbooks(1).Sheets(k).Copy Workbooks(2).Sheets(1)
    ss = ThisWorkbook.Path & "\" & sht.Name & ".xlsx"
    wk.SaveAs ss
    wk.Close
Next
Application.DisplayAlerts = True
MsgBox "��ֹ��������"
End Sub

