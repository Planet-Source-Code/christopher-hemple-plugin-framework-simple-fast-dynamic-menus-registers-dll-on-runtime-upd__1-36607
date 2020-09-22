Attribute VB_Name = "modRegPlug"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0
Private Const ERROR_AHHHHHH = &HF

Public Function RegisterServer(hWnd As Long, DllServerPath As String, bRegister As Boolean)
On Error Resume Next
Dim lb As Long, pa As Long
lb = LoadLibrary(DllServerPath)
If bRegister Then
pa = GetProcAddress(lb, "DllRegisterServer")
Else
pa = GetProcAddress(lb, "DllUnregisterServer")
End If
If CallWindowProc(pa, hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
RegisterServer = ERROR_SUCCESS
Else
RegisterServer = ERROR_AHHHHHH
End If
FreeLibrary lb
End Function
'This works by: File To Regiter,Dll to Open (plugin),and the form that is opening it
Sub LoadPlugin(ByVal RegisterFile As String, ByVal Dll_Open, ByVal Form As Form)
RegisterFile = App.Path + "\" + RegisterFile
RegisterServer Form.hWnd, RegisterFile, True
'Change This Here To Open A Diffrent Sub In The Class Module In The Dll
CreateObject(Dll_Open).Load
RegisterServer Form.hWnd, RegisterFile, False
End Sub

Sub Build_App(Form As Form)
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set Folder = menu.GetFolder(App.Path + "\")
For Each File In Folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile
    If Left(File.Name, 4) = "app_" Then Else GoTo nextfile
    Load Form.mnuAppearance(x)
    With Form.mnuAppearance(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub

Sub Build_Sft(Form As Form)
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set Folder = menu.GetFolder(App.Path + "\")
For Each File In Folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile
    If Left(File.Name, 4) = "sft_" Then Else GoTo nextfile
    Load Form.mnuSoftware(x)
    With Form.mnuSoftware(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub

Sub Build_Reg(Form As Form)
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set Folder = menu.GetFolder(App.Path + "\")
For Each File In Folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile
    If Left(File.Name, 4) = "reg_" Then Else GoTo nextfile
    Load Form.mnuRegistry(x)
    With Form.mnuRegistry(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub

Sub Build_Oth(Form As Form)
On Error Resume Next
x = 0
Dim menu As New FileSystemObject
Set Folder = menu.GetFolder(App.Path + "\")
For Each File In Folder.Files
    If Right(File.Name, 4) = ".dll" Then Else GoTo nextfile
If Left(File.Name, 4) = "app_" Then GoTo nextfile
If Left(File.Name, 4) = "reg_" Then GoTo nextfile
If Left(File.Name, 4) = "sft_" Then GoTo nextfile
    Load Form.mnuOther(x)
    With Form.mnuOther(x)
        .Caption = Mid$(File.Name, 1, Len(File.Name) - 4)
        .Tag = File
        .Visible = True
        .Enabled = True
    End With
    x = x + 1
nextfile:
    Next
End Sub


