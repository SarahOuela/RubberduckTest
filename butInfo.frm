VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} butInfo 
   Caption         =   "Op�ration effectu�e"
   ClientHeight    =   150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2685
   OleObjectBlob   =   "butInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "butInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    'Application.Wait Now + TimeValue("0:00:02")
    Unload butInfo
    DoEvents
End Sub

Private Sub UserForm_Click()

End Sub
