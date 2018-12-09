Private Sub Workbook_Open()
Dim TargetBar As CommandBar
Dim NewMenu As Object
Dim NewItem As Object
Dim NewMenuTemp As Object

Set TargetBar = Application.CommandBars("Worksheet Menu Bar")
TargetBar.Visible = True

For Each NewMenuTemp In TargetBar.Controls
If NewMenuTemp.Caption = "Function" Then
Exit Sub
End If
Next
Set NewMenu = TargetBar.Controls.Add(Type:=msoControlPopup, ID:=1, Temporary:=True)
NewMenu.Caption = "Function"

Set NewItem = NewMenu.Controls.Add(Type:=msoControlButton, ID:=1, Temporary:=True)
NewItem.Caption = "Function 1"
NewItem.OnAction = "模块1.function1"

Exit Sub
End Sub

'以上代码放到thisworkbook里，然后在模块1里添加function1
sub function1()
UserForm1.show
end sub
'这样excel启动后会在excel的菜单上新增个菜单Function，点击里边的Function1就可以了。