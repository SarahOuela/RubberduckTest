VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} selDataSheet 
   Caption         =   "S�lection des feuilles � traiter"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6165
   OleObjectBlob   =   "selDataSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "selDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butSelSheetCancel_Click()
    selDataSheet.Hide
End Sub

Private Sub butSelSheetOK_Click()
    Call DisableExcel
    Call majListDataSheet
    Call EnableExcel
    selDataSheet.Hide
End Sub

Private Sub lvSelInputSheet_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvSelInputSheet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'Dim j As Integer
    If Item.tag = "L" Then
        Item.Checked = True
        msg ("itemBlocked")
    End If
    ''Dim listSheetDataDone() As String
    ''listSheetDataDone = getListSheetDataDone()
    'If Item.Checked = False Then
            'Item.ForeColor = RGB(0, 0, 255) 'Changement couleur
            'Item.Bold = True 'Gras
            'If UBound(listSheetDataDone) > 0 Then
                'For c = LBound(listSheetDataDone) To UBound(listSheetDataDone)
                    'If listSheetDataDone(c) <> "" Then
                        'Item.Checked = True
                        'MsgBox Item.text
                        'msg ("itemBlocked")
                    'End If
                'Next
            'End If

            'For j = 1 To Item.ListSubItems.Count
                'Item.ListSubItems(j).ForeColor = RGB(0, 0, 255)
                'Item.ListSubItems(j).Bold = True
            'Next j
        'Else
            'Item.ForeColor = RGB(1, 0, 0) 'Changement couleur
            'Item.Bold = False
            
            'For j = 1 To Item.ListSubItems.Count
                'Item.ListSubItems(j).ForeColor = RGB(1, 0, 0)
                'Item.ListSubItems(j).Bold = False
            'Next j
    'End If
End Sub

Private Sub UserForm_Activate()
    Dim listSheetData() As String
    Dim listSheetDataPresent() As String
    Dim listSheetDataDone() As Boolean
    listSheetData = getListSheetData("NOP_Col")
    listSheetDataPresent = getListSheetDataPresent()
    listSheetDataDone = getListSheetDataDone()
    lvSelInputSheet.ListItems.Clear
    If UBound(listSheetData) > 0 Then
        For I = LBound(listSheetData) To UBound(listSheetData)
            Set nLine = lvSelInputSheet.ListItems.Add(, "N" & I, listSheetData(I))
            nLine.tag = ""
            If UBound(listSheetDataPresent) > 0 Then
                For j = LBound(listSheetDataPresent) To UBound(listSheetDataPresent)
                    If listSheetData(I) = listSheetDataPresent(j) Then
                        nLine.Checked = True
                        If listSheetDataDone(j) Then
                            nLine.tag = "L"
                        Else
                            nLine.tag = ""
                        End If
                    End If
                Next
            End If
        Next
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With lvSelInputSheet
        With .ColumnHeaders
            .Clear
            .Add , , "Feuille", 305
        End With
    End With
    lvSelInputSheet.View = lvwReport
    lvSelInputSheet.HideColumnHeaders = False
End Sub
