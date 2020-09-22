VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Plugin Framework"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTweaks 
      Caption         =   "Xp Tweaks"
      Begin VB.Menu mnuAppearanceSub 
         Caption         =   "Appearance"
         Begin VB.Menu mnuAppearance 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuSoftwareSub 
         Caption         =   "Software"
         Begin VB.Menu mnuSoftware 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuRegistrySub 
         Caption         =   "Registry"
         Begin VB.Menu mnuRegistry 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuOtherSub 
         Caption         =   "Other"
         Begin VB.Menu mnuOther 
            Caption         =   "(Empty)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnRefreshTweakList 
         Caption         =   "Refresh Tweak List"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To Make a Plugin go into the appearance when your making the project make the projectname "app_" and then a name
'Software menu - same as above but sft_
'Registry Menu - same as above but reg_
'Other menu, dont put any of the about infront of it

'Rules:

'1) the dll name as to be the same as the project name , e.g: project1 --->project1.dll
'2) the dll has to have a class module called main
'3) the dll has to have a sub called load in the class module ( this is the sub it runs )
'4) the dlls have to be in the apps path

Private Sub form_Load()
Build_App Me
Build_Sft Me
Build_Reg Me
Build_Oth Me
End Sub

Private Sub mnRefreshTweakList_Click()
Build_App Me
Build_Sft Me
Build_Reg Me
Build_Oth Me
End Sub

Private Sub mnuAppearance_Click(Index As Integer)
On Error GoTo Err_Handle
LoadPlugin mnuAppearance(Index).Caption + ".dll", mnuAppearance(Index).Caption + ".Main", Me
Exit Sub
Err_Handle:
MsgBox App.Title & " Has Caused A Error And It Will Now Close." & vbCrLf & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, vbCritical, "Error - " & Err.Description
Unload Me
End Sub

Private Sub mnuOther_Click(Index As Integer)
On Error GoTo Err_Handle
LoadPlugin mnuOther(Index).Caption + ".dll", mnuOther(Index).Caption + ".Main", Me
Exit Sub
Err_Handle:
MsgBox App.Title & " Has Caused A Error And It Will Now Close." & vbCrLf & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, vbCritical, "Error - " & Err.Description
Unload Me
End Sub

Private Sub mnuRegistry_Click(Index As Integer)
On Error GoTo Err_Handle
LoadPlugin mnuRegistry(Index).Caption + ".dll", mnuRegistry(Index).Caption + ".Main", Me
Exit Sub
Err_Handle:
MsgBox App.Title & " Has Caused A Error And It Will Now Close." & vbCrLf & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, vbCritical, "Error - " & Err.Description
Unload Me
End Sub

Private Sub mnuSoftware_Click(Index As Integer)
On Error GoTo Err_Handle
LoadPlugin mnuSoftware(Index).Caption + ".dll", mnuSoftware(Index).Caption + ".Main", Me
Exit Sub
Err_Handle:
MsgBox App.Title & " Has Caused A Error And It Will Now Close." & vbCrLf & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Description : " & Err.Description, vbCritical, "Error - " & Err.Description
Unload Me
End Sub

Private Sub mnuTweaks_Click()
Build_App Me
Build_Sft Me
Build_Reg Me
Build_Oth Me
End Sub
