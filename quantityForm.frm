VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} quantityForm 
   Caption         =   "propri�t�s de la quantit�"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16260
   OleObjectBlob   =   "quantityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "quantityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub essai_Click()

End Sub

Private Sub ComboTime_Change()
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
    newtime = Trim(quantityForm.ComboTime.VALUE)
    newqua = Trim(quantityForm.quantity.Caption)
    newGen = Trim(quantityForm.perimetre.Caption)
    Dim lig As String
    For a = LBound(quantity, 1) To UBound(quantity, 1)
        If newtime = time(a, 1) And newqua = qua(a, 1) And newGen = GEN(a, 1) Then
            lig = Trim(LIGNES(a, 1))
            Exit For
        End If
    Next
    'MsgBox quantityForm.ComboTime.Value & Chr(10) & quantityForm.genQua.ListItems(quantityForm.genQua.ListItems.Count).Text
    
    
    If lig = quantityForm.genQua.ListItems(quantityForm.genQua.ListItems.Count).text Then
        ' on ne fait rien
    Else
        ' marche pas bien
        ''''nomenclatureForm.majPanel (lig)
    End If
End Sub


'''Private Sub listGenQua_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '''Dim k As String
    '''Dim ttt As String
    '''If Item.Checked Then
        '''For i = 1 To listGenQua.ListItems.Count
            '''If listGenQua.ListItems(i).Checked Then
                '''If listGenQua.ListItems(i) = Item Then
                    '''k = "" & listGenQua.ListItems(i).key
                    '''nk = "E" & Mid(k, 2)
                    '''lt = nomenclatureForm.analyseTerme(listGenQua.ListItems(i).ListSubItems(nk).Text)
'MsgBox UBound(lt) & ":" & listGenQua.ListItems(i).ListSubItems(nk).Text
                    '''If ListTermes.ListItems.Count <> UBound(lt) Then
                        '''If ListTermes.ListItems.Count > 0 Then
                            'MsgBox ListTermes.ListItems.Count
                            '''For t = 1 To ListTermes.ListItems.Count
                                '''ListTermes.ListItems("t" & t).ListSubItems("h" & t).Text = ""
                            '''Next
                        '''End If
                    '''Else
                        '''If UBound(lt) > 0 Then
                            '''For t = LBound(lt) To UBound(lt)
                                '''ListTermes.ListItems("t" & t).ListSubItems("h" & t).Text = "" & lt(t)
                            '''Next
                        '''End If
                    '''End If
                '''Else
                    '''listGenQua.ListItems(i).Checked = False
                '''End If
            '''End If
        '''Next
    '''Else
        '''If ListTermes.ListItems.Count > 0 Then
            '''For t = 1 To ListTermes.ListItems.Count
                '''ListTermes.ListItems("t" & t).ListSubItems("h" & t).Text = ""
            '''Next
        '''End If
    '''End If
'''End Sub

Private Sub ListTermes_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub UserForm_Click()

End Sub
