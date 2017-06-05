VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufLanguage 
   Caption         =   "Choisissez le langage"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3060
   OleObjectBlob   =   "ufLanguage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()
    Unload ufLanguage
End Sub

Private Sub butCn_Click()
    g_Language_Temp = "cn"
    labCn.Visible = True
    labEn.Visible = False
    labFr.Visible = False
End Sub


Private Sub butEn_Click()
    g_Language_Temp = "en"
    labEn.Visible = True
    labFr.Visible = False
    labCn.Visible = False
End Sub

Private Sub butFr_Click()
    g_Language_Temp = "fr"
    labFr.Visible = True
    labEn.Visible = False
    labCn.Visible = False
End Sub

Private Sub butOK_Click()
    g_Language = g_Language_Temp
    Unload ufLanguage
End Sub

Private Sub UserForm_Initialize()
    labFr.Visible = False
    labEn.Visible = False
    labCn.Visible = False
    If g_Language = "fr" Then
        butFr.SetFocus
        labFr.Visible = True
    End If
    If g_Language = "en" Then
        butEn.SetFocus
        labEn.Visible = True
    End If
    If g_Language = "cn" Then
        butCn.SetFocus
        labCn.Visible = True
    End If
End Sub

