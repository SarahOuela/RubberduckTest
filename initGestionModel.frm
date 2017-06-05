VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} initGestionModel 
   Caption         =   "Gestion du mod�le"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7005
   OleObjectBlob   =   "initGestionModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "initGestionModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Unload initGestionModel
End Sub

Private Sub gestionButton_Click()
    Unload initGestionModel
    nomenclatureVisu.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    cbTimeG.Clear
    cbAreaG.Clear
    cbScenarioG.Clear
    cbNomenclatureG.Clear
    cbQuantityG.Clear
    For Each WS In ThisWorkbook.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        thirdCell = Trim(WS.Cells(1, 3).VALUE)
        fourthCell = Trim(WS.Cells(1, 4).VALUE)
        If firstcell Like "*DATE*" Then
            cbTimeG.AddItem (WS.NAME)
        End If
        If firstcell Like "*AREA*" Then
            cbAreaG.AddItem (WS.NAME)
        End If
        If firstcell Like "*SCENARIO*" Then
            cbScenarioG.AddItem (WS.NAME)
        End If
        If firstcell Like "*Feuille*" And thirdCell Like "*ENTITE*" And fourthCell Like "*SHORTNAME*" Then
            cbNomenclatureG.AddItem (WS.NAME)
        End If
        If firstcell Like "*Feuille*" And thirdCell Like "*ENTITE*" And fourthCell Like "*AREA*" Then
            cbQuantityG.AddItem (WS.NAME)
        End If
    Next WS
    labTimeG.Caption = "Temps"
    labAreaG.Caption = "Zones"
    labScenarioG.Caption = "Sc�narios"
    labNomenclatureG.Caption = "Nomenclature"
    labQuantityG.Caption = "Quantit�s"
    cbTimeG.ListIndex = 0
    cbAreaG.ListIndex = 0
    cbScenarioG.ListIndex = 0
    cbNomenclatureG.ListIndex = 0
    cbQuantityG.ListIndex = 0
End Sub
