VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nomenclatureVisu 
   Caption         =   "Gestion du mod�le"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18960
   OleObjectBlob   =   "nomenclatureVisu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "nomenclatureVisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nomListNomChe() As String
Dim nomListNomAttChe() As String
Dim quaListNomAttChe() As String
Dim nomListNomGenChe() As String
Dim NOMENCLATURE() As Variant
Dim quantity() As Variant
Dim LIGNES() As Variant
Dim ORIGINE() As Variant
Dim FLNOM As Worksheet
Dim FLQUA As Worksheet
'ReDim nomListNomChe(0)
Private Sub CommandButton1_Click()
    '''Set FL = Worksheets("Technique")
    '''nomenclatureVisu.Caption = FL.Cells(1, 1)
End Sub

Private Sub dimensions_Change()
    'MsgBox "ici"
    'if Me.dimensions.Pages(Me.dimensions.VALUE).Caption = "Temps"
    'If Me.dimensions.VALUE = 1 Then
    'MsgBox Me.dimensions.Pages(Me.dimensions.VALUE).Caption
    If Me.dimensions.VALUE = 3 Then Call majNomenclature
End Sub

Private Sub dimensions_Enter()
    'MsgBox "oui"
End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub listNomenclature_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub nomGenLabelAll_Click()

End Sub

Private Sub treeNomenclatureQ_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Activate()
    With Me
        .StartUpPosition = 3
        .Width = Application.Width - 40
        .Height = Application.Height - 40
        .Left = 20
        .Top = 20
    End With
    ratiow = Application.Width / Me.Width
    ratioh = Application.Height / Me.Height
    
    ' les tab
    Me.Controls("dimensions").Left = 0
    Me.Controls("dimensions").Top = 30
    Me.Controls("dimensions").Width = Me.Width - 4
    Me.Controls("dimensions").Height = Me.Height - Me.Controls("dimensions").Top - 21
    
    Me.Controls("treeAttribut").Height = 100
    Me.Controls("treeAreaAttribut").Height = 100
    Me.Controls("treeScenarioAttribut").Height = 100
    Me.Controls("treeAttribut").Top = Me.Controls("dimensions").Height - Me.Controls("treeAttribut").Height - 15
    Me.Controls("treeAreaAttribut").Top = Me.Controls("dimensions").Height - Me.Controls("treeAreaAttribut").Height - 15
    Me.Controls("treeScenarioAttribut").Top = Me.Controls("dimensions").Height - Me.Controls("treeScenarioAttribut").Height - 15
    Me.Controls("treeAttributLabel").Top = Me.Controls("treeAttribut").Top - 15
    Me.Controls("treeAreaAttributLabel").Top = Me.Controls("treeAreaAttribut").Top - 15
    Me.Controls("treeScenarioAttributLabel").Top = Me.Controls("treeScenarioAttribut").Top - 15
    Me.Controls("treeNomenclature").Height = Me.Controls("treeAttributLabel").Top - Me.Controls("treeNomenclature").Top
    Me.Controls("treeArea").Height = Me.Controls("treeAreaAttributLabel").Top - Me.Controls("treeArea").Top
    Me.Controls("treeScenario").Height = Me.Controls("treeScenarioAttributLabel").Top - Me.Controls("treeScenario").Top
    Me.Controls("listNomenclature").Height = Me.Controls("dimensions").Height / 2
    Me.Controls("listNomenclature").Top = Me.Controls("dimensions").Height - Me.Controls("listNomenclature").Height - 15
    Me.Controls("listNomenclature").Width = Me.Controls("dimensions").Width - Me.Controls("treeAttribut").Width
    Me.Controls("listNomenclatureLabel").Top = Me.Controls("listNomenclature").Top - Me.Controls("listNomenclatureLabel").Height
    'nomLabelAll, nomLabel, nomRet
    Me.Controls("nomLabelAll").Top = Me.Controls("listNomenclatureLabel").Top + 3
    Me.Controls("nomLabel").Top = Me.Controls("listNomenclatureLabel").Top + 3
    Me.Controls("nomRet").Top = Me.Controls("listNomenclatureLabel").Top + 3
    'Me.Controls("listNomenclatureGen").Top = Me.Controls("listNomenclatureLabel").Top + Me.Controls("listNomenclatureLabel").Height
    'MsgBox Me.Controls("listNomenclatureLabel").Top & ":" & Me.Controls("listNomenclatureGen").Top
    Me.Controls("listNomenclatureGen").Height = Me.Controls("listNomenclatureLabel").Top - Me.Controls("listNomenclatureGen").Top
    Me.Controls("listNomenclatureGen").Width = Me.Controls("dimensions").Width - Me.Controls("treeAttribut").Width
    Me.Controls("nomGenLabelAll").Top = Me.Controls("listNomenclatureGenLabel").Top + 3
    Me.Controls("nomGenLabel").Top = Me.Controls("listNomenclatureGenLabel").Top + 3
    Me.Controls("nomGenRet").Top = Me.Controls("listNomenclatureGenLabel").Top + 3
    
    ' le tab quantity
    Me.Controls("treeAreaAttrQ").Height = 100
    Me.Controls("treeScenarioAttrQ").Height = 100
    Me.Controls("treeAttributQ").Height = 100
    Me.Controls("treeAreaAttrQ").Top = Me.Controls("dimensions").Height - Me.Controls("treeAttributQ").Height - 15
    Me.Controls("treeScenarioAttrQ").Top = Me.Controls("treeAreaAttrQ").Top - Me.Controls("treeScenarioAttrQ").Height
    Me.Controls("treeAttributQ").Top = Me.Controls("treeScenarioAttrQ").Top - Me.Controls("treeAttributQ").Height
End Sub
Private Sub listQuantityGen_DblClick()
    Dim QUANTITYALL() As Variant
    Set FLQUA = Worksheets("Quantit�s Model")
    derlignom = FLQUA.Range("C" & FLQUA.Rows.Count).End(xlUp).Row
    QUANTITYALL = FLQUA.Range("A1:AP" & derlignom).VALUE
    qua = QUANTITYALL(CInt(listQuantityGen.SelectedItem), 34)
    Lib = QUANTITYALL(CInt(listQuantityGen.SelectedItem), 31)
    quantityForm.Caption = "Propri�t�s de la quantit� : " & qua
    quantityForm.listGenQua.View = lvwReport
    quantityForm.listGenQua.HideColumnHeaders = False
    quantityForm.genQua.View = lvwReport
    quantityForm.genQua.HideColumnHeaders = False
    quantityForm.ListTermes.View = lvwReport
    quantityForm.ListTermes.HideColumnHeaders = False
    With quantityForm.ListTermes
        With .ColumnHeaders
            .Clear
            '''If maxFacteur > 0 Then
            .Add , , "Termes g�n�riques", 300
            .Add , , "Termes �tendus", 500
            '''For i = 1 To maxFacteur
            '''.Add , , "hi�rarchie " & i, 80
            '''.Add , , "entit� " & i, 50
            '''.Add , , "attribut " & i, 40
            '''Next
            '''End If
        End With
    End With
    Call filtrerQuantityFromDbl(listQuantityGen.SelectedItem)
    'Set nLine = quantityForm.listGenQua.ListItems.Add(, "N" & 1, "" & 1)
    quantityForm.Show
End Sub
Function analyseTerme(perO As String) As String()
    Dim res() As String
    ReDim res(0)
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = True
    reg.Pattern = "(\[.*?\].*?\))"
    If reg.test(perO) Then
        Set matches = reg.Execute(perO)
        For Each Match In matches
            If UBound(res) > 0 Then ReDim Preserve res(1 To (UBound(res) + 1))
            If UBound(res) = 0 Then ReDim res(1 To 1)
            res(UBound(res)) = Match
        Next
    End If
    analyseTerme = res
End Function
Private Sub listNomenclatureGen_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call listNomenclatureGen_ItemCheck_Core
    '''Call filtrerNomenclatureFromGen(listCheckedNodes)
End Sub
Private Sub listNomenclatureGen_ItemCheck_Core()
    Dim listCheckedNodes() As String
    ReDim listCheckedNodes(0)
    For I = 1 To treeNomenclature.Nodes.Count
        If treeNomenclature.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeNomenclature.Nodes(I).key
        End If
    Next
    Dim listCheckedNodesAtt() As String
    ReDim listCheckedNodesAtt(0)
    For I = 1 To treeAttribut.Nodes.Count
        If treeAttribut.Nodes(I).Checked Then
            If UBound(listCheckedNodesAtt) = 0 Then
                ReDim listCheckedNodesAtt(1 To 1)
            Else
                ReDim Preserve listCheckedNodesAtt(1 To UBound(listCheckedNodesAtt) + 1)
            End If
            listCheckedNodesAtt(UBound(listCheckedNodesAtt)) = treeAttribut.Nodes(I).key
        End If
    Next
    Call filtrerNomenclatureAttributFromGen(listCheckedNodes, listCheckedNodesAtt)
End Sub

Private Sub filtrerNomenclatureAttributFromGen(listCheckedNodes() As String, listCheckedNodesAtt() As String)
    Dim nomList() As String
    ReDim nomList(1 To UBound(NOMENCLATURE, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    listNomenclature.ListItems.Clear
    Dim bool As Boolean
    Dim listNomGen() As String
    ReDim listNomGen(0)
    Dim maxFacteur As Integer
    maxFacteur = 1
    ReDim nomListNomGenChe(0)
    For I = 1 To listNomenclatureGen.ListItems.Count
        If listNomenclatureGen.ListItems(I).Checked Then
            If UBound(nomListNomGenChe) = 0 Then
                ReDim nomListNomGenChe(1 To 1)
            Else
                ReDim Preserve nomListNomGenChe(1 To UBound(nomListNomGenChe) + 1)
            End If
            nomListNomGenChe(UBound(nomListNomGenChe)) = listNomenclatureGen.ListItems(I).key
        End If
    Next
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        nomList(a) = NOMENCLATURE(a, 1)
        bool = True
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
            where = ">" & Replace(nomList(a), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
        Next
        If bool Then
            where = Replace(where, ":", ",")
            where = Replace(where, ">", ",>") & ","
            For at = LBound(listCheckedNodesAtt) To UBound(listCheckedNodesAtt)
                If Trim(listCheckedNodesAtt(at)) <> "" Then bool = bool And (where Like "*," & listCheckedNodesAtt(at) & ",*")
            Next
        End If
        If bool Then maxFacteur = WorksheetFunction.Max(UBound(Split(nomList(a), ".")) + 1, maxFacteur)
    Next
    Dim listKey() As String
    Dim ec As Integer
    ec = 0
    If listNomenclatureGen.ListItems.Count = 0 Then
        Exit Sub
    End If
    ReDim listKey(1 To listNomenclatureGen.ListItems.Count)
    For I = 1 To listNomenclatureGen.ListItems.Count
        listKey(I) = listNomenclatureGen.ListItems(I).key
        If listNomenclatureGen.ListItems(I).Checked Then ec = ec + 1
    Next
    Dim jj As Integer
    jj = 0
    Dim kk As Integer
    kk = 0
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        nomList(a) = NOMENCLATURE(a, 1)
        bool = True
        kk = kk + 1
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
            where = ">" & Replace(nomList(a), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
            If bool Then
                where = Replace(where, ":", ",")
                where = Replace(where, ">", ",>") & ","
                For at = LBound(listCheckedNodesAtt) To UBound(listCheckedNodesAtt)
                    If Trim(listCheckedNodesAtt(at)) <> "" Then bool = bool And (where Like "*," & listCheckedNodesAtt(at) & ",*")
                Next
            End If
            If Not bool Then
                Exit For
            Else
                jj = jj + 1
                spll = Split(LIGNES(a, 1), ",")
                boolOr = False
                For j = LBound(spll) To UBound(spll)
                    n = CInt(spll(j))
                    bools = False
                    For k = 1 To UBound(listKey)
                        If listKey(k) = "N" & n Then
                            bools = True
                            Exit For
                        End If
                    Next
                    If bools Then
                        boolOr = boolOr Or listNomenclatureGen.ListItems("N" & n).Checked
                    Else
                        boolO = False
                    End If
                Next
                bool = bool And boolOr
            End If
        Next
        If bool Then
            ii = ii + 1
            If UBound(listNomGen) = 0 Then
                ReDim listNomGen(1 To 1)
            Else
                ReDim Preserve listNomGen(1 To UBound(listNomGen) + 1)
            End If
            listNomGen(UBound(listNomGen)) = LIGNES(a, 1)
            Set nLine = listNomenclature.ListItems.Add(, "N" & a, "+")
            Set nLineG = nLine.ListSubItems.Add(, "G" & a, "" & a)
            Set nLineL = nLine.ListSubItems.Add(, "L" & a, LIGNES(a, 1))
            splP = Split(nomList(a), ".")
            If splP(0) Like "*>*" Then
                strr = StrReverse(splP(0))
                splitr0 = StrReverse(Split(strr, ">")(0))
                par = Mid(splP(0), 1, Len(splP(0)) - Len(splitr0) - 1)
                splitA = Split(splitr0 & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, par)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, avat)
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            Else
                splitA = Split(splP(0) & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, avat)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, "")
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            End If
            If UBound(splP) > 0 Then
                For I = LBound(splP) + 1 To UBound(splP)
                    j = I + 1
                    If splP(I) Like "*>*" Then
                        strr = StrReverse(splP(I))
                        splitr0 = StrReverse(Split(strr, ">")(0))
                        par = Mid(splP(I), 1, Len(splP(I)) - Len(splitr0) - 1)
                        splitA = Split(splitr0 & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, par)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, avat)
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    Else
                        splitA = Split(splP(I) & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, avat)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, "")
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    End If
                Next
            End If
        End If
    Next
    '''nomLabel.Caption = ii & " entit�s �tendues sur " & jj & " parmi " & kk
    nomLabel.Caption = jj & " entit�s retenues"
    nomGenRet.Caption = ec & " entit�s coch�es"
    nomRet.Caption = ii & " entit�s coch�es"
    '''nomGenLabel.Caption = ig & " entit�s retenues"
End Sub
Private Sub filtrerQuantityFromGen(listCheckedNodes() As String)

    Dim LIGNES() As Variant
    Dim ORIGINE() As Variant
    Dim AREA() As Variant
    Dim SCENARIO() As Variant
    Dim qua() As Variant
    Dim time() As Variant
    Dim nom() As Variant
    Dim equation() As Variant
    Dim nomList() As String
    Set FLQUA = Worksheets(initGestionModel.cbQuantityG.VALUE)
    derlignom = FLQUA.Range("A" & FLQUA.Rows.Count).End(xlUp).Row
    quantity = FLQUA.Range("C2:C" & derlignom).VALUE
    LIGNES = FLQUA.Range("B2:B" & derlignom).VALUE
    ORIGINE = FLQUA.Range("A2:A" & derlignom).VALUE
    AREA = FLQUA.Range("D2:D" & derlignom).VALUE
    SCENARIO = FLQUA.Range("E2:E" & derlignom).VALUE
    qua = FLQUA.Range("F2:F" & derlignom).VALUE
    time = FLQUA.Range("G2:G" & derlignom).VALUE
    nom = FLQUA.Range("K2:K" & derlignom).VALUE
    equation = FLQUA.Range("H2:H" & derlignom).VALUE
    ReDim nomList(1 To UBound(quantity, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    listQuantity.ListItems.Clear
    Dim bool As Boolean
    Dim listNomGen() As String
    ReDim listNomGen(0)

    Dim ec As Integer
    ec = 0
    Dim listKey() As String
    ReDim listKey(1 To listQuantityGen.ListItems.Count)
    For I = 1 To listQuantityGen.ListItems.Count
        listKey(I) = listQuantityGen.ListItems(I).key
        If listQuantityGen.ListItems(I).Checked Then ec = ec + 1
    Next
    Dim jj As Integer
    jj = 0
    Dim kk As Integer
    kk = 0
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        nomList(a) = quantity(a, 1)
        bool = True
        kk = kk + 1
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
where = ">" & Replace(nomList(a), ".", ">") & ">" & Replace(AREA(a, 1), ".", ">") & ">" & Replace(SCENARIO(a, 1), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
            
            If Not bool Then
                Exit For
            Else
                jj = jj + 1
                spll = Split(LIGNES(a, 1), ",")
                boolOr = False
                For j = LBound(spll) To UBound(spll)
                    n = CInt(spll(j))
                    bools = False
                    For k = 1 To UBound(listKey)
                        If listKey(k) = "N" & n Then
                            bools = True
                            Exit For
                        End If
                    Next
                    If bools Then
                        boolOr = boolOr Or listQuantityGen.ListItems("N" & n).Checked
                    Else
                        boolO = False
                    End If
                Next
                bool = bool And boolOr
            End If
        Next
        If bool Then
            ii = ii + 1
            If UBound(listNomGen) = 0 Then
                ReDim listNomGen(1 To 1)
            Else
                ReDim Preserve listNomGen(1 To UBound(listNomGen) + 1)
            End If
            listNomGen(UBound(listNomGen)) = LIGNES(a, 1)
            Set nLine = listQuantity.ListItems.Add(, "N" & a, "" & a)
            Set nLineL = nLine.ListSubItems.Add(, "L" & a, LIGNES(a, 1))
            Set nLineF = nLine.ListSubItems.Add(, "F" & a, qua(a, 1))
            Set nLineG = nLine.ListSubItems.Add(, "G" & a, time(a, 1))
            Set nLineN = nLine.ListSubItems.Add(, "N" & a, nom(a, 1))
            Set nLinee = nLine.ListSubItems.Add(, "E" & a, equation(a, 1))
            Set nLinea = nLine.ListSubItems.Add(, "a" & a, AREA(a, 1))
            Set nLines = nLine.ListSubItems.Add(, "s" & a, SCENARIO(a, 1))
            splP = Split(nomList(a), ".")
            If splP(0) Like "*>*" Then
                strr = StrReverse(splP(0))
                splitr0 = StrReverse(Split(strr, ">")(0))
                par = Mid(splP(0), 1, Len(splP(0)) - Len(splitr0) - 1)
                splitA = Split(splitr0 & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, par)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, avat)
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            Else
                splitA = Split(splP(0) & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, avat)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, "")
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            End If
            If UBound(splP) > 0 Then
                For I = LBound(splP) + 1 To UBound(splP)
                    j = I + 1
                    If splP(I) Like "*>*" Then
                        strr = StrReverse(splP(I))
                        splitr0 = StrReverse(Split(strr, ">")(0))
                        par = Mid(splP(I), 1, Len(splP(I)) - Len(splitr0) - 1)
                        splitA = Split(splitr0 & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, par)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, avat)
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    Else
                        splitA = Split(splP(I) & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, avat)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, "")
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    End If
                Next
            End If
        End If
    Next
    quaLabel.Caption = jj & " quantit�s retenues"
    quaGenRet.Caption = ec & " quantit�s coch�es"
    quaRet.Caption = ii & " quantit�s coch�es"

    '''quaLabel.Caption = ii & " Quantit�s �tendues sur " & jj & " parmi " & kk
    '''quaGenLabel.Caption = ig & " quantit�s g�n�riques parmi "
    listQuantity.View = lvwReport
    listQuantity.HideColumnHeaders = False
End Sub
Function majPanel(key As String)
    Call filtrerQuantityFromDbl(key)
End Function
Private Sub filtrerQuantityFromDbl(itemNode As String)
    Dim quantity() As Variant
    Dim LIGNES() As Variant
    Dim ORIGINE() As Variant
    Dim AREA() As Variant
    Dim SCENARIO() As Variant
    Dim GEN() As Variant
    Dim qua() As Variant
    Dim time() As Variant
    Dim equation() As Variant
    Dim nomList() As String
    Dim FLQUA As Worksheet
    Set FLQUA = Worksheets(initGestionModel.cbQuantityG.VALUE)
    derlignom = FLQUA.Range("A" & FLQUA.Rows.Count).End(xlUp).Row
    quantity = FLQUA.Range("C2:C" & derlignom).VALUE
    LIGNES = FLQUA.Range("B2:B" & derlignom).VALUE
    ORIGINE = FLQUA.Range("A2:A" & derlignom).VALUE
    AREA = FLQUA.Range("D2:D" & derlignom).VALUE
    SCENARIO = FLQUA.Range("E2:E" & derlignom).VALUE
    GEN = FLQUA.Range("J2:J" & derlignom).VALUE
    qua = FLQUA.Range("F2:F" & derlignom).VALUE
    time = FLQUA.Range("G2:G" & derlignom).VALUE
    equation = FLQUA.Range("H2:H" & derlignom).VALUE
    ReDim nomList(1 To UBound(quantity, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    listQuantity.ListItems.Clear
    Dim bool As Boolean
    Dim listNomGen() As String
    ReDim listNomGen(0)
    Dim maxFacteur As Integer
    Dim FLGEN As Worksheet
    Set FLGEN = Worksheets(ORIGINE(1, 1))
    derlignom = FLGEN.Range("C" & FLGEN.Rows.Count).End(xlUp).Row
    Dim GENDATA() As Variant
    GENDATA = FLGEN.Range("A1:AL" & derlignom).VALUE
    
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        If InStr("," & LIGNES(a, 1) & ",", "," & itemNode & ",") > 0 Then
            maxFacteur = UBound(Split(quantity(a, 1), ".")) + 1
            Exit For
        End If
    Next
    litsGenTerms = analyseTerme(Trim(GENDATA(CInt(itemNode), 38)))
    If UBound(litsGenTerms) > 0 Then
        For t = LBound(litsGenTerms) To UBound(litsGenTerms)
            Set nLinet = quantityForm.ListTermes.ListItems.Add(, "t" & t, litsGenTerms(t))
            Set nLineh = nLinet.ListSubItems.Add(, "h" & t, "")
        Next
    End If
    With quantityForm.genQua
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
                .Add , , "GENERIQUE", 50
                .Add , , "Equation", 150
                .Add , , "area", 30
                .Add , , "sc�nario", 50
                For I = 1 To maxFacteur
                    .Add , , "hi�rarchie " & I, 80
                    .Add , , "entit� " & I, 50
                    .Add , , "attribut " & I, 40
                Next
            End If
        End With
    End With
    'quantityForm.time.Caption = " " & Trim(GENDATA(CInt(itemNode), 37))
    quantityForm.equation.Caption = " " & Trim(GENDATA(CInt(itemNode), 38))
    quantityForm.quantity.Caption = " " & Trim(GENDATA(CInt(itemNode), 34))
    quantityForm.quantityName.Caption = " " & Trim(GENDATA(CInt(itemNode), 31))
    Set nLineG = quantityForm.genQua.ListItems.Add(, "N" & itemNode, itemNode)
    Set nLinee = nLineG.ListSubItems.Add(, "E" & itemNode, GENDATA(CInt(itemNode), 38))
    Set nLinea = nLineG.ListSubItems.Add(, "a" & itemNode, GENDATA(CInt(itemNode), 32))
    Set nLines = nLineG.ListSubItems.Add(, "s" & itemNode, GENDATA(CInt(itemNode), 33))
    For I = 1 To maxFacteur
        Set nLineh = nLineG.ListSubItems.Add(, I & "H" & itemNode, GENDATA(CInt(itemNode), 3 * I + 0))
        Set nLinet = nLineG.ListSubItems.Add(, I & "T" & itemNode, GENDATA(CInt(itemNode), 3 * I + 1))
        Set nLinea = nLineG.ListSubItems.Add(, I & "A" & itemNode, GENDATA(CInt(itemNode), 3 * I + 2))
    Next
    With quantityForm.listGenQua
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
                .Add , , "ETENDU", 50
                .Add , , "Equation", 150
                .Add , , "area", 30
                .Add , , "sc�nario", 50
                For I = 1 To maxFacteur
                    .Add , , "hi�rarchie " & I, 80
                    .Add , , "entit� " & I, 50
                    .Add , , "attribut " & I, 40
                Next
            End If
        End With
    End With
    Dim perim As String
    Dim quaId As String
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        If InStr("," & LIGNES(a, 1) & ",", "," & itemNode & ",") > 0 Then
            quantityForm.perimetre.Caption = " " & Trim(GEN(a, 1))
            perim = Trim(GEN(a, 1))
            quaId = Trim(qua(a, 1))
            Exit For
        End If
    Next
    Dim listPerQua() As String
    ReDim listPerQua(0)
    Dim bb As Boolean
    Dim li As Integer
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        If Trim(GEN(a, 1)) = perim And quaId = qua(a, 1) Then
            If UBound(listPerQua) = 0 Then
                ReDim listPerQua(1 To 1)
                listPerQua(UBound(listPerQua)) = Trim(time(a, 1)) & ":" & Trim(LIGNES(a, 1))
                quantityForm.ComboTime.AddItem (time(a, 1))
                If InStr("," & LIGNES(a, 1) & ",", "," & itemNode & ",") > 0 Then li = quantityForm.ComboTime.ListCount - 1
            Else
                bb = False
                For lpq = LBound(listPerQua) To UBound(listPerQua)
                    If listPerQua(lpq) = Trim(time(a, 1)) & ":" & Trim(LIGNES(a, 1)) Then
                        bb = True
                        Exit For
                    End If
                Next
                If Not bb Then
                    ReDim Preserve listPerQua(1 To UBound(listPerQua) + 1)
                    listPerQua(UBound(listPerQua)) = Trim(time(a, 1)) & ":" & Trim(LIGNES(a, 1))
                    quantityForm.ComboTime.AddItem (time(a, 1))
                    If InStr("," & LIGNES(a, 1) & ",", "," & itemNode & ",") > 0 Then li = quantityForm.ComboTime.ListCount - 1
                End If
            End If
        End If
        If InStr("," & LIGNES(a, 1) & ",", "," & itemNode & ",") > 0 Then
            Set nLine = quantityForm.listGenQua.ListItems.Add(, "N" & a, "" & a)
            Set nLinee = nLine.ListSubItems.Add(, "E" & a, equation(a, 1))
            Set nLinea = nLine.ListSubItems.Add(, "a" & a, AREA(a, 1))
            Set nLines = nLine.ListSubItems.Add(, "s" & a, SCENARIO(a, 1))
            splP = Split(quantity(a, 1), ".")
            If splP(0) Like "*>*" Then
                strr = StrReverse(splP(0))
                splitr0 = StrReverse(Split(strr, ">")(0))
                par = Mid(splP(0), 1, Len(splP(0)) - Len(splitr0) - 1)
                splitA = Split(splitr0 & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, par)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, avat)
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            Else
                splitA = Split(splP(0) & ":", ":")
                avat = splitA(0)
                apat = splitA(1)
                Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, avat)
                Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, "")
                Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            End If
            If UBound(splP) > 0 Then
                For I = LBound(splP) + 1 To UBound(splP)
                    j = I + 1
                    If splP(I) Like "*>*" Then
                        strr = StrReverse(splP(I))
                        splitr0 = StrReverse(Split(strr, ">")(0))
                        par = Mid(splP(I), 1, Len(splP(I)) - Len(splitr0) - 1)
                        splitA = Split(splitr0 & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, par)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, avat)
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    Else
                        splitA = Split(splP(I) & ":", ":")
                        avat = splitA(0)
                        apat = splitA(1)
                        Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, avat)
                        Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, "")
                        Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    End If
                Next
            End If
        End If
    Next
    quantityForm.ComboTime.ListIndex = li
    quantityForm.listGenQua.View = lvwReport
    quantityForm.listGenQua.HideColumnHeaders = False
    quantityForm.genQua.View = lvwReport
    quantityForm.genQua.HideColumnHeaders = False
End Sub



Private Sub MultiPage1_Change()

End Sub

'''Private Sub listQuantityGen_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    'For j = 1 To listQuantity.ListItems.Count
        'listQuantity.ListItems(j).Checked = False
        'listQuantity.ListItems(j).ForeColor = vbBlack
        'listQuantity.ListItems(j).Bold = False
    'Next
    
    'For i = 1 To listQuantityGen.ListItems.Count
        'If listQuantityGen.ListItems(i).Checked Then
            'For j = 1 To listQuantity.ListItems.Count
                'If InStr("," & listQuantity.ListItems(j).ListSubItems(1).Text & ",", "," & listQuantityGen.ListItems(i).Text & ",") > 0 Then
                    'listQuantity.ListItems(j).Checked = True
                    'listQuantity.ListItems(j).ForeColor = vbRed
                    'listQuantity.ListItems(j).Bold = True
                    'aaa = listQuantity.ListItems(j)
                'End If
            'Next
        'End If
    'Next
    'listQuantity.SelectedItem = aaa
    'listQuantity.SelectedItem.EnsureVisible
    '''Dim listCheckedNodes() As String
    '''ReDim listCheckedNodes(0)
    '''For i = 1 To treeNomQ.Nodes.Count
        '''If treeNomQ.Nodes(i).Checked Then
            '''If UBound(listCheckedNodes) = 0 Then
                '''ReDim listCheckedNodes(1 To 1)
            '''Else
                '''ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            '''End If
            '''listCheckedNodes(UBound(listCheckedNodes)) = treeNomQ.Nodes(i).key
        '''End If
    '''Next
    '''For i = 1 To treeAreaQ.Nodes.Count
        '''If treeAreaQ.Nodes(i).Checked Then
            '''If UBound(listCheckedNodes) = 0 Then
                '''ReDim listCheckedNodes(1 To 1)
            '''Else
                '''ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            '''End If
            '''listCheckedNodes(UBound(listCheckedNodes)) = treeAreaQ.Nodes(i).key
        '''End If
    '''Next
    '''For i = 1 To treeScenarioQ.Nodes.Count
        '''If treeScenarioQ.Nodes(i).Checked Then
            '''If UBound(listCheckedNodes) = 0 Then
                '''ReDim listCheckedNodes(1 To 1)
            '''Else
                '''ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            '''End If
            '''listCheckedNodes(UBound(listCheckedNodes)) = treeScenarioQ.Nodes(i).key
        '''End If
    '''Next
    '''Call filtrerQuantityFromGen(listCheckedNodes)
'''End Sub

Private Sub quitButton_Click()
    Unload Me
End Sub

Private Sub treeAttribut_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub treeAttribut_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim listCheckedNodes() As String
    ReDim listCheckedNodes(0)
    For I = 1 To treeNomenclature.Nodes.Count
        If treeNomenclature.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeNomenclature.Nodes(I).key
        End If
    Next
    Dim listCheckedNodesAtt() As String
    ReDim listCheckedNodesAtt(0)
    For I = 1 To treeAttribut.Nodes.Count
        If treeAttribut.Nodes(I).Checked Then
            If UBound(listCheckedNodesAtt) = 0 Then
                ReDim listCheckedNodesAtt(1 To 1)
            Else
                ReDim Preserve listCheckedNodesAtt(1 To UBound(listCheckedNodesAtt) + 1)
            End If
            listCheckedNodesAtt(UBound(listCheckedNodesAtt)) = treeAttribut.Nodes(I).key
        End If
    Next
    nomListNomChe = listCheckedNodes
    nomListNomAttChe = listCheckedNodesAtt
    Call filtrerNomenclatureAttribut(listCheckedNodes, listCheckedNodesAtt, 1)
End Sub

Private Sub treeNomenclature_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim listCheckedNodes() As String
    ReDim listCheckedNodes(0)
    For I = 1 To treeNomenclature.Nodes.Count
        If treeNomenclature.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeNomenclature.Nodes(I).key
        End If
    Next
    Dim listCheckedNodesAtt() As String
    ReDim listCheckedNodesAtt(0)
    For I = 1 To treeAttribut.Nodes.Count
        If treeAttribut.Nodes(I).Checked Then
            If UBound(listCheckedNodesAtt) = 0 Then
                ReDim listCheckedNodesAtt(1 To 1)
            Else
                ReDim Preserve listCheckedNodesAtt(1 To UBound(listCheckedNodesAtt) + 1)
            End If
            listCheckedNodesAtt(UBound(listCheckedNodesAtt)) = treeAttribut.Nodes(I).key
        End If
    Next
    nomListNomChe = listCheckedNodes
    nomListNomAttChe = listCheckedNodesAtt
    Call filtrerNomenclatureAttribut(listCheckedNodes, listCheckedNodesAtt, 1)
End Sub
Private Sub filtrerNomenclatureAttribut(listCheckedNodes() As String, listCheckedNodesAtt() As String, init As Integer)
    Dim nomList() As String
    ReDim nomList(1 To UBound(NOMENCLATURE, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    listNomenclature.ListItems.Clear
    Dim bool As Boolean
    Dim listNomGen() As String
    ReDim listNomGen(0)
    Dim maxFacteur As Integer
    maxFacteur = 1
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        nomList(a) = NOMENCLATURE(a, 1)
        bool = True
        where = ">" & Replace(nomList(a), ".", ">")
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
        Next
        If bool Then
            where = Replace(where, ":", ",")
            where = Replace(where, ">", ",>") & ","
            For at = LBound(listCheckedNodesAtt) To UBound(listCheckedNodesAtt)
                If Trim(listCheckedNodesAtt(at)) <> "" Then bool = bool And (where Like "*," & listCheckedNodesAtt(at) & ",*")
            Next
        End If
        If bool Then maxFacteur = WorksheetFunction.Max(UBound(Split(nomList(a), ".")) + 1, maxFacteur)
    Next
    Dim al As Integer
    al = 0
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        nomList(a) = NOMENCLATURE(a, 1)
        bool = True
        al = al + 1
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
            where = ">" & Replace(nomList(a), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
        Next
        If bool Then
            where = Replace(where, ":", ",")
            where = Replace(where, ">", ",>") & ","
            For at = LBound(listCheckedNodesAtt) To UBound(listCheckedNodesAtt)
                If Trim(listCheckedNodesAtt(at)) <> "" Then bool = bool And (where Like "*," & listCheckedNodesAtt(at) & ",*")
            Next
        End If
        If bool Then
            ii = ii + 1
            If UBound(listNomGen) = 0 Then
                ReDim listNomGen(1 To 1)
            Else
                ReDim Preserve listNomGen(1 To UBound(listNomGen) + 1)
            End If
            listNomGen(UBound(listNomGen)) = LIGNES(a, 1)
        End If
    Next
    listGen = Split(Join(listNomGen, ","), ",")
    Dim lastlistchecked() As String
    ReDim lastlistchecked(0)
    For I = 1 To listNomenclatureGen.ListItems.Count
        If listNomenclatureGen.ListItems(I).Checked Then
            lastlistchecked = addItemToEnd(listNomenclatureGen.ListItems(I).key, lastlistchecked)
        End If
    Next
    listNomenclatureGen.ListItems.Clear
    Dim FLGEN As Worksheet
    Set FLGEN = Worksheets(ORIGINE(1, 1))
    derlignom = FLGEN.Range("C" & FLGEN.Rows.Count).End(xlUp).Row
    Dim GENDATA() As Variant
    GENDATA = FLGEN.Range("A1:AE" & derlignom).VALUE
    Dim keyUsed() As String
    ReDim keyUsed(0)
    Dim dejaPresent As Boolean
    Dim ig As Integer
    ig = 0
    For I = LBound(listGen) To UBound(listGen)
        If I = LBound(listGen) Then
            ig = 1
            ReDim keyUsed(1 To 1)
            keyUsed(1) = listGen(I)
            Set nLine = listNomenclatureGen.ListItems.Add(, "N" & listGen(I), "+")
            Set nLineG = nLine.ListSubItems.Add(, "G" & listGen(I), "" & listGen(I))
            For c = LBound(lastlistchecked) To UBound(lastlistchecked)
                If lastlistchecked(c) = "N" & listGen(I) Then
                    nLine.Checked = True
                    Exit For
                End If
            Next
            Set nLinee = nLine.ListSubItems.Add(, "E" & listGen(I), GENDATA(CInt(listGen(I)), 30))
            For j = 1 To maxFacteur
                Set nLineh = nLine.ListSubItems.Add(, "h" & j & "h" & I, GENDATA(CInt(listGen(I)), j * 3 + 0))
                Set nLinee = nLine.ListSubItems.Add(, "e" & j & "e" & I, GENDATA(CInt(listGen(I)), j * 3 + 1))
                Set nLinea = nLine.ListSubItems.Add(, "a" & j & "a" & I, GENDATA(CInt(listGen(I)), j * 3 + 2))
            Next
        End If
        dejaPresent = False
        For k = LBound(keyUsed) To UBound(keyUsed)
            If keyUsed(k) = listGen(I) Then
                ' d�j� pr�sent on en fait rien
                dejaPresent = True
                Exit For
            End If
        Next
        If Not dejaPresent Then
            ReDim Preserve keyUsed(1 To UBound(keyUsed) + 1)
            keyUsed(UBound(keyUsed)) = listGen(I)
            Set nLine = listNomenclatureGen.ListItems.Add(, "N" & listGen(I), "+")
            Set nLineG = nLine.ListSubItems.Add(, "G" & listGen(I), "" & listGen(I))
            For c = LBound(lastlistchecked) To UBound(lastlistchecked)
                If lastlistchecked(c) = "N" & listGen(I) Then
                    nLine.Checked = True
                    Exit For
                End If
            Next
            Set nLinee = nLine.ListSubItems.Add(, "E" & listGen(I), GENDATA(CInt(listGen(I)), 30))
            ig = ig + 1
            For j = 1 To maxFacteur
                Set nLineh = nLine.ListSubItems.Add(, "h" & j & "h" & I, GENDATA(CInt(listGen(I)), j * 3 + 0))
                Set nLinee = nLine.ListSubItems.Add(, "e" & j & "e" & I, GENDATA(CInt(listGen(I)), j * 3 + 1))
                Set nLinea = nLine.ListSubItems.Add(, "a" & j & "a" & I, GENDATA(CInt(listGen(I)), j * 3 + 2))
            Next
        End If
    Next
    nomLabel.Caption = ii & " entit�s retenues"
    
    If init = 0 Then nomLabelAll.Caption = al & " entit�s �tendues"
    If init = 0 Then nomGenLabelAll.Caption = ig & " entit�s g�n�riques"
    nomGenRet.Caption = "0 entit� coch�e"
    nomRet.Caption = "0 entit� coch�e"
    nomGenLabel.Caption = ig & " entit�s retenues"
    'listNomenclature.View = lvwReport
    'listNomenclature.HideColumnHeaders = False
    Call listNomenclatureGen_ItemCheck_Core
End Sub
Private Sub filtrerQuantity(listCheckedNodes() As String, init As Integer)
    Dim quantity() As Variant
    Dim LIGNES() As Variant
    Dim ORIGINE() As Variant
    Dim AREA() As Variant
    Dim SCENARIO() As Variant
    Dim nom() As Variant
    Dim nomList() As String
    Dim FLQUA As Worksheet
    Set FLQUA = Worksheets(initGestionModel.cbQuantityG.VALUE)
    derlignom = FLQUA.Range("A" & FLQUA.Rows.Count).End(xlUp).Row
    quantity = FLQUA.Range("C2:C" & derlignom).VALUE
    LIGNES = FLQUA.Range("B2:B" & derlignom).VALUE
    ORIGINE = FLQUA.Range("A2:A" & derlignom).VALUE
    AREA = FLQUA.Range("D2:D" & derlignom).VALUE
    SCENARIO = FLQUA.Range("E2:E" & derlignom).VALUE
    nom = FLQUA.Range("K2:K" & derlignom).VALUE
    ReDim nomList(1 To UBound(quantity, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    listQuantity.ListItems.Clear
    Dim bool As Boolean
    Dim listNomGen() As String
    ReDim listNomGen(0)
    Dim maxFacteur As Integer
    maxFacteur = 1
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        nomList(a) = quantity(a, 1)
        bool = True
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
where = ">" & Replace(nomList(a), ".", ">") & ">" & Replace(AREA(a, 1), ".", ">") & ">" & Replace(SCENARIO(a, 1), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
        Next
        If bool Then maxFacteur = WorksheetFunction.Max(UBound(Split(nomList(a), ".")) + 1, maxFacteur)
    Next
    With listQuantityGen
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
            .Add , , "N� lig", 35
            .Add , , "Nb Ext", 40
            .Add , , "Quantit�", 50
            .Add , , "Time", 40
            .Add , , "Nom", 60
            .Add , , "Equation", 100
            .Add , , "area", 30
            .Add , , "sc�nario", 50
            For I = 1 To maxFacteur
                .Add , , "hi�rarchie " & I, 80
                .Add , , "entit� " & I, 50
                .Add , , "attribut " & I, 40
            Next
            End If
        End With
    End With
    Dim kk As Integer
    kk = 0
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        nomList(a) = quantity(a, 1)
        kk = kk + 1
        bool = True
        For I = LBound(listCheckedNodes) To UBound(listCheckedNodes)
where = ">" & Replace(nomList(a), ".", ">") & ">" & Replace(AREA(a, 1), ".", ">") & ">" & Replace(SCENARIO(a, 1), ".", ">")
            bool = bool And (where Like "*>" & listCheckedNodes(I) & "*")
        Next
        If bool Then
            ii = ii + 1
            If UBound(listNomGen) = 0 Then
                ReDim listNomGen(1 To 1)
            Else
                ReDim Preserve listNomGen(1 To UBound(listNomGen) + 1)
            End If
            listNomGen(UBound(listNomGen)) = LIGNES(a, 1)
            Set nLine = listQuantity.ListItems.Add(, "N" & a, "" & a)
            Set nLineL = nLine.ListSubItems.Add(, "L" & a, LIGNES(a, 1))
            Set nLinea = nLine.ListSubItems.Add(, "a" & a, AREA(a, 1))
            Set nLines = nLine.ListSubItems.Add(, "s" & a, SCENARIO(a, 1))
            '''splP = Split(nomList(a), ".")
            '''If splP(0) Like "*>*" Then
                '''strr = StrReverse(splP(0))
                '''splitr0 = StrReverse(Split(strr, ">")(0))
                '''par = Mid(splP(0), 1, Len(splP(0)) - Len(splitr0) - 1)
                '''splitA = Split(splitr0 & ":", ":")
                '''avat = splitA(0)
                '''apat = splitA(1)
                '''Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, par)
                '''Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, avat)
                '''Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            '''Else
                '''splitA = Split(splP(0) & ":", ":")
                '''avat = splitA(0)
                '''apat = splitA(1)
                '''Set nLine0 = nLine.ListSubItems.Add(, "h1h" & a, avat)
                '''Set nLine1 = nLine.ListSubItems.Add(, "e1e" & a, "")
                '''Set nLine2 = nLine.ListSubItems.Add(, "a1a" & a, apat)
            '''End If
            '''If UBound(splP) > 0 Then
                '''For i = LBound(splP) + 1 To UBound(splP)
                    '''j = i + 1
                    '''If splP(i) Like "*>*" Then
                        '''strr = StrReverse(splP(i))
                        '''splitr0 = StrReverse(Split(strr, ">")(0))
                        '''par = Mid(splP(i), 1, Len(splP(i)) - Len(splitr0) - 1)
                        '''splitA = Split(splitr0 & ":", ":")
                        '''avat = splitA(0)
                        '''apat = splitA(1)
                        '''Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, par)
                        '''Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, avat)
                        '''Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    '''Else
                        '''splitA = Split(splP(i) & ":", ":")
                        '''avat = splitA(0)
                        '''apat = splitA(1)
                        '''Set nLine0 = nLine.ListSubItems.Add(, "h" & j & "h" & a, avat)
                        '''Set nLine1 = nLine.ListSubItems.Add(, "e" & j & "e" & a, "")
                        '''Set nLine2 = nLine.ListSubItems.Add(, "a" & j & "a" & a, apat)
                    '''End If
                '''Next
            '''End If
        End If
    Next
    listQuantity.View = lvwReport
    listQuantity.HideColumnHeaders = False
    quaLabel.Caption = " 0 Quantit� �tendue sur " & ii & " parmi " & kk
    listGen = Split(Join(listNomGen, ","), ",")
    listQuantityGen.ListItems.Clear
    Dim FLGEN As Worksheet
    Set FLGEN = Worksheets(ORIGINE(1, 1))
    derlignom = FLGEN.Range("C" & FLGEN.Rows.Count).End(xlUp).Row
    Dim GENDATA() As Variant
    Dim QUAG() As Variant
    Dim TIMEG() As Variant
    Dim NOMG() As Variant
    Dim EQUATIONG() As Variant
    Dim AREAG() As Variant
    Dim SCENARIOG() As Variant
    GENDATA = FLGEN.Range("A1:AE" & derlignom).VALUE
    QUAG = FLGEN.Range("AH1:AH" & derlignom).VALUE
    TIMEG = FLGEN.Range("AK1:AK" & derlignom).VALUE
    NOMG = FLGEN.Range("AE1:AE" & derlignom).VALUE
    EQUATIONG = FLGEN.Range("AL1:AL" & derlignom).VALUE
    AREAG = FLGEN.Range("AF1:AF" & derlignom).VALUE
    SCENARIOG = FLGEN.Range("AG1:AG" & derlignom).VALUE
    Dim keyUsed() As String
    ReDim keyUsed(0)
    Dim dejaPresent As Boolean
    Dim ig As Integer
    ig = 0
    Dim al As Integer
    al = 0
    For I = LBound(listGen) To UBound(listGen)
        If I = LBound(listGen) Then
            ig = 1
            ReDim keyUsed(1 To 1)
            keyUsed(1) = listGen(I)
            Set nLine = listQuantityGen.ListItems.Add(, "N" & listGen(I), "" & listGen(I))
            Set nLinee = nLine.ListSubItems.Add(, "E" & listGen(I), GENDATA(CInt(listGen(I)), 30))
            Set nLineF = nLine.ListSubItems.Add(, "F" & listGen(I), QUAG(CInt(listGen(I)), 1))
            Set nLineG = nLine.ListSubItems.Add(, "G" & listGen(I), TIMEG(CInt(listGen(I)), 1))
            Set nLineN = nLine.ListSubItems.Add(, "N" & listGen(I), NOMG(CInt(listGen(I)), 1))
            Set nLineK = nLine.ListSubItems.Add(, "K" & listGen(I), EQUATIONG(CInt(listGen(I)), 1))
            Set nLinea = nLine.ListSubItems.Add(, "A" & listGen(I), AREAG(CInt(listGen(I)), 1))
            Set nLines = nLine.ListSubItems.Add(, "S" & listGen(I), SCENARIOG(CInt(listGen(I)), 1))
            For j = 1 To maxFacteur
                Set nLineh = nLine.ListSubItems.Add(, "h" & j & "h" & I, GENDATA(CInt(listGen(I)), j * 3 + 0))
                Set nLinee = nLine.ListSubItems.Add(, "e" & j & "e" & I, GENDATA(CInt(listGen(I)), j * 3 + 1))
                Set nLinea = nLine.ListSubItems.Add(, "a" & j & "a" & I, GENDATA(CInt(listGen(I)), j * 3 + 2))
            Next
        End If
        dejaPresent = False
        For k = LBound(keyUsed) To UBound(keyUsed)
            If keyUsed(k) = listGen(I) Then
                ' d�j� pr�sent on en fait rien
                dejaPresent = True
                Exit For
            End If
        Next
        If Not dejaPresent Then
            ReDim Preserve keyUsed(1 To UBound(keyUsed) + 1)
            keyUsed(UBound(keyUsed)) = listGen(I)
            Set nLine = listQuantityGen.ListItems.Add(, "N" & listGen(I), "" & listGen(I))
            Set nLinee = nLine.ListSubItems.Add(, "E" & listGen(I), GENDATA(CInt(listGen(I)), 30))
            Set nLineF = nLine.ListSubItems.Add(, "F" & listGen(I), QUAG(CInt(listGen(I)), 1))
            Set nLineG = nLine.ListSubItems.Add(, "G" & listGen(I), TIMEG(CInt(listGen(I)), 1))
            Set nLineN = nLine.ListSubItems.Add(, "N" & listGen(I), NOMG(CInt(listGen(I)), 1))
            Set nLineK = nLine.ListSubItems.Add(, "K" & listGen(I), EQUATIONG(CInt(listGen(I)), 1))
            Set nLinea = nLine.ListSubItems.Add(, "A" & listGen(I), AREAG(CInt(listGen(I)), 1))
            Set nLines = nLine.ListSubItems.Add(, "S" & listGen(I), SCENARIOG(CInt(listGen(I)), 1))
            ig = ig + 1
            For j = 1 To maxFacteur
                Set nLineh = nLine.ListSubItems.Add(, "h" & j & "h" & I, GENDATA(CInt(listGen(I)), j * 3 + 0))
                Set nLinee = nLine.ListSubItems.Add(, "e" & j & "e" & I, GENDATA(CInt(listGen(I)), j * 3 + 1))
                Set nLinea = nLine.ListSubItems.Add(, "a" & j & "a" & I, GENDATA(CInt(listGen(I)), j * 3 + 2))
            Next
        End If
    Next
    quaLabel.Caption = ii & " quantit�s retenues"
    If init = 0 Then quaLabelAll.Caption = kk & " quantit�s �tendues"
    If init = 0 Then quaGenLabelAll.Caption = ig & " quantit�s g�n�riques"
    quaGenRet.Caption = "0 quantit� coch�e"
    quaRet.Caption = "0 quantit� coch�e"
    quaGenLabel.Caption = ig & " quantit�s retenues"
    '''quaGenLabel.Caption = ig & " quantit�s g�n�riques parmi " & al
    listQuantityGen.View = lvwReport
    listQuantityGen.HideColumnHeaders = False
End Sub



'''Private Sub treeNomQ_NodeCheck(ByVal Node As MSComctlLib.Node)
    '''Call callFiltrerQuantity
'''End Sub
'''Private Sub treeAreaQ_NodeCheck(ByVal Node As MSComctlLib.Node)
    '''Call callFiltrerQuantity
'''End Sub
'''Private Sub treeScenarioQ_NodeCheck(ByVal Node As MSComctlLib.Node)
    '''Call callFiltrerQuantity
'''End Sub
Private Sub callFiltrerQuantity()
    Dim listCheckedNodes() As String
    ReDim listCheckedNodes(0)
    For I = 1 To treeNomenclatureQ.Nodes.Count
        If treeNomenclatureQ.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeNomenclatureQ.Nodes(I).key
        End If
    Next
    For I = 1 To treeAreaQ.Nodes.Count
        If treeAreaQ.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeAreaQ.Nodes(I).key
        End If
    Next
    For I = 1 To treeScenarioQ.Nodes.Count
        If treeScenarioQ.Nodes(I).Checked Then
            If UBound(listCheckedNodes) = 0 Then
                ReDim listCheckedNodes(1 To 1)
            Else
                ReDim Preserve listCheckedNodes(1 To UBound(listCheckedNodes) + 1)
            End If
            listCheckedNodes(UBound(listCheckedNodes)) = treeScenarioQ.Nodes(I).key
        End If
    Next
    Call filtrerQuantity(listCheckedNodes, 1)
End Sub

Private Sub treeNomQ_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub treeNomenclatureQ_NodeCheck(ByVal Node As MSComctlLib.Node)
    callFiltrerQuantity
End Sub

Private Sub treeArea_BeforeLabelEdit(Cancel As Integer)

End Sub
Private Sub initScenario()
    ' lecture des scenarios
    Dim SCENARIO() As Variant
    Dim scenList() As String
    Dim FLSCEN As Worksheet
    Set FLSCEN = Worksheets(initGestionModel.cbScenarioG.VALUE)
    derlignom = FLSCEN.Range("A" & FLSCEN.Rows.Count).End(xlUp).Row
    SCENARIO = FLSCEN.Range("A2:A" & derlignom).VALUE
    ReDim scenList(1 To UBound(SCENARIO, 1))
    Dim listScenKey() As String
    ReDim listScenKey(0)
    ii = 0
    For a = LBound(SCENARIO, 1) To UBound(SCENARIO, 1)
        scenList(a) = SCENARIO(a, 1)
        If InStr(scenList(a), ".") = 0 Then
            If Not scenList(a) Like "*>*" Then
                ' cas de la racine
                Set nParent1 = treeScenario.Nodes.Add(key:=scenList(a), text:=scenList(a))
                Set nParent2 = treeScenarioQ.Nodes.Add(key:=scenList(a), text:=scenList(a))
                If UBound(listScenKey) = 0 Then
                    ReDim listScenKey(1 To 1)
                Else
                    ReDim Preserve listScenKey(1 To UBound(listScenKey) + 1)
                End If
                listScenKey(UBound(listScenKey)) = "@" & scenList(a) & "@"
            Else
                strr = StrReverse(scenList(a))
                splitr0 = StrReverse(Split(strr, ">")(0))
                splitrr = StrReverse(splitr0)
                par = Mid(scenList(a), 1, Len(scenList(a)) - Len(splitrr) - 1)
                If UBound(Filter(listScenKey, "@" & par & "@", True)) >= 0 Then
                    Set nParent1 = treeScenario.Nodes.Item(par)
                    Set nParent2 = treeScenarioQ.Nodes.Item(par)
                    Set nChild1 = treeScenario.Nodes.Add(nParent1, tvwChild, scenList(a), splitr0)
                    Set nChild2 = treeScenarioQ.Nodes.Add(nParent2, tvwChild, scenList(a), splitr0)
                    If UBound(listScenKey) = 0 Then
                        ReDim listScenKey(1 To 1)
                    Else
                        ReDim Preserve listScenKey(1 To UBound(listScenKey) + 1)
                    End If
                    listScenKey(UBound(listScenKey)) = "@" & scenList(a) & "@"
                End If
            End If
        End If
    Next
    'listQuantity.CheckBoxes = True
    'listQuantityGen.CheckBoxes = True
    treeScenario.CheckBoxes = True
    treeScenario.Nodes("s").Expanded = True
    treeScenarioQ.CheckBoxes = True
    treeScenarioQ.Nodes("s").Expanded = True
    treeArea.Nodes("a").Expanded = True
    treeAreaQ.Nodes("a").Expanded = True
End Sub
Private Sub initNomenclature()
    ' Initialisation des listes
    Dim Item As String
    Dim maxFacteur As Integer
    maxFacteur = 1
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        Item = NOMENCLATURE(a, 1)
        maxFacteur = WorksheetFunction.Max(UBound(Split(Item, ".")) + 1, maxFacteur)
    Next
    With listNomenclatureGen
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
                .Add , , "", 20
                .Add , , "N� lig", 35
                .Add , , "Nb", 40
                For I = 1 To maxFacteur
                    .Add , , "hi�rarchie " & I, 80
                    .Add , , "entit� " & I, 50
                    .Add , , "attribut " & I, 40
                Next
            End If
        End With
    End With
    With listNomenclature
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
                .Add , , "", 20
                .Add , , "N� lig", 35
                .Add , , "From", 40
                For I = 1 To maxFacteur
                    .Add , , "hi�rarchie " & I, 80
                    .Add , , "entit� " & I, 50
                    .Add , , "attribut " & I, 40
                Next
            End If
        End With
    End With
    listNomenclatureGen.View = lvwReport
    listNomenclatureGen.HideColumnHeaders = False
    listNomenclature.View = lvwReport
    listNomenclature.HideColumnHeaders = False
End Sub
Private Sub initQuantity()
    ' Initialisation des listes
    Dim Item As String
    Dim maxFacteur As Integer
    maxFacteur = 1
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        Item = quantity(a, 1)
        maxFacteur = WorksheetFunction.Max(UBound(Split(Item, ".")) + 1, maxFacteur)
    Next
    With listQuantityGen
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
            .Add , , "N� lig", 35
            .Add , , "Nb Ext", 40
            .Add , , "Quantit�", 50
            .Add , , "Time", 40
            .Add , , "Nom", 60
            .Add , , "Equation", 100
            .Add , , "area", 30
            .Add , , "sc�nario", 50
            For I = 1 To maxFacteur
                .Add , , "hi�rarchie " & I, 80
                .Add , , "entit� " & I, 50
                .Add , , "attribut " & I, 40
            Next
            End If
        End With
    End With
    With listQuantity
        With .ColumnHeaders
            .Clear
            If maxFacteur > 0 Then
            .Add , , "N� lig", 40
            .Add , , "From", 40
            .Add , , "Quantit�", 50
            .Add , , "Time", 40
            .Add , , "Nom", 60
            .Add , , "Equation", 100
            .Add , , "area", 30
            .Add , , "sc�nario", 50
            For I = 1 To maxFacteur
                .Add , , "hi�rarchie " & I, 80
                .Add , , "entit� " & I, 50
                .Add , , "attribut " & I, 40
            Next
            End If
        End With
    End With
    listQuantityGen.View = lvwReport
    listQuantityGen.HideColumnHeaders = False
    listQuantity.View = lvwReport
    listQuantity.HideColumnHeaders = False
End Sub
Private Sub majNomenclature()
    If UBound(nomListNomChe) > 0 Then
        For I = 1 To treeNomenclature.Nodes.Count
            For j = LBound(nomListNomChe) To UBound(nomListNomChe)
                If treeNomenclature.Nodes(I).key = nomListNomChe(j) Then treeNomenclature.Nodes(I).Checked = True
            Next
        Next
    End If
    If UBound(nomListNomAttChe) > 0 Then
        For I = 1 To treeAttribut.Nodes.Count
            For j = LBound(nomListNomAttChe) To UBound(nomListNomAttChe)
                If treeAttribut.Nodes(I).key = nomListNomAttChe(j) Then treeAttribut.Nodes(I).Checked = True
            Next
        Next
    End If
    For I = 1 To listNomenclatureGen.ListItems.Count
        listNomenclatureGen.ListItems(I).Checked = False
        If UBound(nomListNomGenChe) > 0 Then
            For j = LBound(nomListNomGenChe) To UBound(nomListNomGenChe)
                If listNomenclatureGen.ListItems(I).key = nomListNomGenChe(j) Then listNomenclatureGen.ListItems(I).Checked = True
            Next
        End If
    Next
End Sub
Private Sub UserForm_Initialize()
    ' lecture de la nomenclature
    Me.Caption = "Gestion du mod�le : ( Area=" & initGestionModel.cbAreaG.VALUE & " , Scenario=" & initGestionModel.cbScenarioG.VALUE & " , Time=" & initGestionModel.cbTimeG.VALUE & " , Nomenclature=" & initGestionModel.cbNomenclatureG.VALUE & " , Quantity=" & initGestionModel.cbQuantityG.VALUE & " )"
    '''Dim NOMENCLATURE() As Variant
    Dim nomList() As String
    '''Dim FLNOM As Worksheet
    Set FLNOM = Worksheets(initGestionModel.cbNomenclatureG.VALUE)
    derlignom = FLNOM.Range("A" & FLNOM.Rows.Count).End(xlUp).Row
    NOMENCLATURE = FLNOM.Range("C2:C" & derlignom).VALUE
    LIGNES = FLNOM.Range("B2:B" & derlignom).VALUE
    ORIGINE = FLNOM.Range("A2:A" & derlignom).VALUE
    ' lecture des quantit�s
    Set FLQUA = Worksheets(initGestionModel.cbNomenclatureG.VALUE)
    derlignom = FLQUA.Range("A" & FLQUA.Rows.Count).End(xlUp).Row
    quantity = FLQUA.Range("C2:C" & derlignom).VALUE
    ReDim nomList(1 To UBound(NOMENCLATURE, 1))
    Dim listNomKey() As String
    ReDim listNomKey(0)
    ii = 0
    'listNomenclature.CheckBoxes = True
    listNomenclatureGen.CheckBoxes = True
    listQuantityGen.CheckBoxes = True
    Dim listNomAttr() As String
    ReDim listNomAttr(0)
    Dim iteman As String
    For a = LBound(NOMENCLATURE, 1) To UBound(NOMENCLATURE, 1)
        nomList(a) = NOMENCLATURE(a, 1)
        If InStr(nomList(a), ".") = 0 Then
            If InStr(nomList(a), ":") > 0 Then
                splAtt = Split(Split(nomList(a), ":")(1), ",")
                For sa = LBound(splAtt) To UBound(splAtt)
                    iteman = splAtt(sa)
                    oldlistNomAttr = listNomAttr
                    listNomAttr = addItemToList(iteman, listNomAttr)
                    If UBound(listNomAttr) <> UBound(oldlistNomAttr) Then
                        Set nParentAtt = treeAttribut.Nodes.Add(key:=iteman, text:=iteman)
                        Set nParentAtt = treeAttributQ.Nodes.Add(key:=iteman, text:=iteman)
                    End If
                Next
            End If
            If Not nomList(a) Like "*>*" Then
                ' cas de la racine
'If nomList(a) = "SECTEUR>INDUSTRIE>BRANCHE>INDUSTRIE_MANUFACTURIERE>Chimie:mixte" Then MsgBox nomList(a)
                Set nParent1 = treeNomenclature.Nodes.Add(key:=nomList(a), text:=nomList(a))
                Set nParent2 = treeNomenclatureQ.Nodes.Add(key:=nomList(a), text:=nomList(a))
                If UBound(listNomKey) = 0 Then
                    ReDim listNomKey(1 To 1)
                Else
                    ReDim Preserve listNomKey(1 To UBound(listNomKey) + 1)
                End If
                listNomKey(UBound(listNomKey)) = "@" & nomList(a) & "@"
            Else
                strr = StrReverse(nomList(a))
                splitr0 = StrReverse(Split(strr, ">")(0))
                splitrr = StrReverse(splitr0)
                par = Mid(nomList(a), 1, Len(nomList(a)) - Len(splitrr) - 1)
                listRes = Filter(listNomKey, "@" & par, True)
                If UBound(listRes) >= 0 Then
                    test = False
                    
                    For R = LBound(listRes) To UBound(listRes)
                        splr = Split(listRes(R), ":")(0)
                        If splr = "@" & par & "@" Or splr = "@" & par Then
                            cle = Mid(listRes(R), 2)
                            cle = Mid(cle, 1, Len(cle) - 1)
                            test = True
                            Exit For
                        End If
                    Next
                    If test Then
                        Set nParent1 = treeNomenclature.Nodes.Item(cle)
                        Set nParent2 = treeNomenclatureQ.Nodes.Item(cle)
                        Set nChild1 = treeNomenclature.Nodes.Add(nParent1, tvwChild, nomList(a), splitr0)
                        Set nChild2 = treeNomenclatureQ.Nodes.Add(nParent2, tvwChild, nomList(a), splitr0)
                        If UBound(listNomKey) = 0 Then
                            ReDim listNomKey(1 To 1)
                        Else
                            ReDim Preserve listNomKey(1 To UBound(listNomKey) + 1)
                        End If
                        listNomKey(UBound(listNomKey)) = "@" & nomList(a) & "@"
                    End If
                End If
            End If
        End If
    Next
    'MsgBox initGestionModel.cbNomenclatureG.VALUE
'Exit Sub
    'For Each nod In treeNomenclature.Nodes
        'If Not nod.Expanded Then
            'nod.Expanded = True
        'End If
    'Next
    '''treeNomenclature.Nodes("k0").Expanded = True
    ' lecture des zone
    Dim AREA() As Variant
    Dim areaList() As String
    Dim FLAREA As Worksheet
    Set FLAREA = Worksheets(initGestionModel.cbAreaG.VALUE)
    derlignom = FLAREA.Range("A" & FLAREA.Rows.Count).End(xlUp).Row
    If derlignom < 3 Then derlignom = 3
    AREA = FLAREA.Range("A2:A" & derlignom).VALUE
    ReDim areaList(1 To UBound(AREA, 1))
    Dim listAreaKey() As String
    ReDim listAreaKey(0)
    Dim listAreaAttr() As String
    ReDim listAreaAttr(0)
    ii = 0
    Dim itema As String
    For a = LBound(AREA, 1) To UBound(AREA, 1)
        areaList(a) = AREA(a, 1)
        If Trim(areaList(a)) <> "" Then
            If InStr(areaList(a), ".") = 0 Then
                If InStr(areaList(a), ":") > 0 Then
                    splAtt = Split(Split(areaList(a), ":")(1), ",")
                    For sa = LBound(splAtt) To UBound(splAtt)
                        itema = splAtt(sa)
                        oldlistAreaAttr = listAreaAttr
                        listAreaAttr = addItemToList(itema, listAreaAttr)
                        If UBound(listAreaAttr) <> UBound(oldlistAreaAttr) Then
                            Set nParentAtt = treeAreaAttribut.Nodes.Add(key:=itema, text:=itema)
                        End If
                    Next
                End If
                If Not areaList(a) Like "*>*" Then
                    ' cas de la racine
                    Set nParent1 = treeArea.Nodes.Add(key:=areaList(a), text:=areaList(a))
                    Set nParent2 = treeAreaQ.Nodes.Add(key:=areaList(a), text:=areaList(a))
                    If UBound(listAreaKey) = 0 Then
                        ReDim listAreaKey(1 To 1)
                    Else
                        ReDim Preserve listAreaKey(1 To UBound(listAreaKey) + 1)
                    End If
                    listAreaKey(UBound(listAreaKey)) = "@" & areaList(a) & "@"
                Else
                    strr = StrReverse(areaList(a))
                    splitr0 = StrReverse(Split(strr, ">")(0))
                    splitrr = StrReverse(splitr0)
                    par = Mid(areaList(a), 1, Len(areaList(a)) - Len(splitrr) - 1)
                    If UBound(Filter(listAreaKey, "@" & par & "@", True)) >= 0 Then
                        Set nParent1 = treeArea.Nodes.Item(par)
                        Set nParent2 = treeAreaQ.Nodes.Item(par)
                        Set nChild1 = treeArea.Nodes.Add(nParent1, tvwChild, areaList(a), splitr0)
                        Set nChild2 = treeAreaQ.Nodes.Add(nParent2, tvwChild, areaList(a), splitr0)
                        If UBound(listAreaKey) = 0 Then
                            ReDim listAreaKey(1 To 1)
                        Else
                            ReDim Preserve listAreaKey(1 To UBound(listAreaKey) + 1)
                        End If
                        listAreaKey(UBound(listAreaKey)) = "@" & areaList(a) & "@"
                    End If
                End If
            End If
        End If
    Next
    '''treeArea.Nodes("a").Expanded = True
    Call initScenario
    Call initNomenclature
    Call initQuantity
    'listNomenclature.View = lvwReport
    'listNomenclature.HideColumnHeaders = False
    'listNomenclatureGen.View = lvwReport
    'listNomenclatureGen.HideColumnHeaders = False
    'listNomenclatureGen.CheckBoxes = True
    Dim listCheckedNodes() As String
    ReDim listCheckedNodes(0)
    Call filtrerNomenclatureAttribut(listCheckedNodes, listCheckedNodes, 0)
    'Call filtrerQuantity(listCheckedNodes, 0)
    ' positionnement sur la page nomenclature
    ReDim nomListNomChe(0)
    ReDim nomListNomAttChe(0)
    ReDim nomListNomGenChe(0)
    Me.dimensions.VALUE = 3
End Sub
