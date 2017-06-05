VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} selInputSheet 
   Caption         =   "S�lection des feuilles de donn�es"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6165
   OleObjectBlob   =   "selInputSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "selInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub butSelSheetCancel_Click()
    selInputSheet.Hide
End Sub

Private Sub butSelSheetOK_Click()
    Call DisableExcel
    Call majListInputSheet
    Call EnableExcel
    selInputSheet.Hide
End Sub

Private Sub lvSelInputSheet_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim listSheetData() As String
    listSheetData = getListSheetData("INPUT")
    lvSelInputSheet.ListItems.Clear
    If UBound(listSheetData) > 0 Then
        For I = LBound(listSheetData) To UBound(listSheetData)
            Set nLine = lvSelInputSheet.ListItems.Add(, "N" & I, listSheetData(I))
        Next
    End If
End Sub
