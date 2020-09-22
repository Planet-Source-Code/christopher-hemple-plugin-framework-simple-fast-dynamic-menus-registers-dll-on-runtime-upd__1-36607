Attribute VB_Name = "Module1"
Sub Build_App(Form As Form)
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set folder = menu.GetFolder(App.Path + "\")
For Each File In folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile
    If Left(File.Name, 4) = "app_" Then Else GoTo nextfile
    Load Form.mnuShowPlugins(x)
    With Form.mnuShowPlugins(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub
