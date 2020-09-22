VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'RULES:
'The Plugins/Dlls Need To Be In the Apps Path. If You Want To Change This refer To The Module
'The dll's Name has to be the same as the project name
'e.g project1 -----> Project1.dll ( Other Wise It WONT Work)
'The Dlls Class Has To Have A Sub Called Load, this is where it will start the code
'the Dlls Class Has To Be Called Main
'the dll's Extention has to be in lower case
'
'This is the sub that the host will try to open
'By Chris Hemple
'Chris The_vb_don@hotmail.com

Sub Load()
MsgBox "The Template Is Working"
End Sub
