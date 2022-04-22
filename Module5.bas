Attribute VB_Name = "Module4"
Public Sub ShortCut(Position As String, SourePath As String) 'Create Link
Set WshShell = CreateObject("Wscript.shell")
With WshShell.CreateShortcut(Position) 'Link Path
If sIcon = "" Then
.Iconlocation = SourePath 'Link Icon
Else
.Iconlocation = sIcon 'Link Icon
End If
.TargetPath = SourePath 'Soure File
.Hotkey = "" 'Hotkey
.save
End With
Set WshShell = Nothing 'Close WshShell Object
End Sub
