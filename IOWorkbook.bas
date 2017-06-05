Attribute VB_Name = "IOWorkbook"
Public Function setWS(nam As String, wb As Workbook) As Worksheet
    If Not WsExist(nam, wb) Then
        wb.Worksheets.Add.Move After:=wb.Worksheets(wb.Worksheets.Count)
        wb.Worksheets(wb.Worksheets.Count).NAME = nam
    End If
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set setWS = wb.Worksheets(nam)
End Function
Function WsExist(nom As String, wb As Workbook) As Boolean
    g_WSerror = nom
    On Error Resume Next
    WsExist = wb.Sheets(nom).Index
End Function
Public Sub setLanguageButtonVisible()
    ThisWorkbook.Worksheets(g_CONTROL).butLangageFr.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).butLangageEn.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).butLangageCn.Visible = False
    If g_Language = "fr" Then ThisWorkbook.Worksheets(g_CONTROL).butLangageFr.Visible = True
    If g_Language = "en" Then ThisWorkbook.Worksheets(g_CONTROL).butLangageEn.Visible = True
    If g_Language = "cn" Then ThisWorkbook.Worksheets(g_CONTROL).butLangageCn.Visible = True
    ThisWorkbook.Worksheets(g_Language_Sheet).Cells(1, 1).VALUE = g_Language
End Sub
Function getLineFrom(sheet2process As String, FLCTL As Worksheet, lig As Integer, col As Integer) As Integer
    getLineFrom = 0
    For I = 1 To 99
        If FLCTL.Cells(lig + I - 1, col).VALUE = sheet2process Then
            getLineFrom = lig + I - 1
            Exit Function
        End If
    Next
End Function
Function WsExistInWB(nom As String, wb As Workbook) As Boolean
    'For i = 1 To wb.Worksheets.Count
        'MsgBox wb.Worksheets(i).NAME
    'Next i
    On Error Resume Next
    WsExistInWB = wb.Worksheets(nom).Index
End Function
''' Retourne la liste des Sheet du modèle dont la cellule (1,1) contient l'argument
Function getListSheetData(ByVal mot As String) As String()
    Dim b As Boolean
    b = openFileIfNot("MODELE")
    Dim res() As String
    Dim firstcell As String
    ReDim res(0)
    For Each WS In g_WB_Modele.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        If firstcell Like "*" & mot & "*" Then
            res = addToList(res, WS.NAME)
        End If
    Next WS
    getListSheetData = res
End Function
''' Retourne la liste des Sheet déjà présents
Function getListSheetDataPresent() As String()
    Dim res() As String
    Dim plage() As Variant
    ReDim res(0)
    plage = ActiveWorkbook.Worksheets(g_CONTROL).Range(DecAlph(g_CONTROL_DATA_C) & g_CONTROL_DATA_L & ":" & DecAlph(g_CONTROL_DATA_C) & (g_CONTROL_DATA_L + 19)).VALUE
    For Nolig = LBound(plage, 1) To UBound(plage, 1)
        If Trim(plage(Nolig, 1)) <> "" Then
            res = addToList(res, Trim(plage(Nolig, 1)))
        End If
    Next
    getListSheetDataPresent = res
End Function
''' Retourne la liste des Sheet déjà générés parmi les présents
Function getListSheetDataDone() As Boolean()
    Dim res() As Boolean
    Dim plage() As Variant
    ReDim res(0)
    plage = ActiveWorkbook.Worksheets(g_CONTROL).Range(DecAlph(g_CONTROL_DATA_C) & g_CONTROL_DATA_L & ":" & DecAlph(g_CONTROL_DATA_C + 1) & (g_CONTROL_DATA_L + 19)).VALUE
    For Nolig = LBound(plage, 1) To UBound(plage, 1)
        If Trim(plage(Nolig, 1)) <> "" Then
            res = addToListBool(res, False)
            If Trim(plage(Nolig, 2)) <> "" Then res(UBound(res)) = True
        End If
    Next
    getListSheetDataDone = res
End Function
Sub majDataName(numcb As Integer, checOrNot As Boolean, ligne As Integer, colonne As Integer)
    ' Mettre en Bold ou non selon checOrNot
    'ThisWorkbook.Worksheets(g_CONTROL).Cells(g_CONTROL_DATA_L + numcb - 1, g_CONTROL_DATA_C).Font.Bold = checOrNot
    ThisWorkbook.Worksheets(g_CONTROL).Cells(ligne + numcb - 1, colonne).Font.Bold = checOrNot
    'If checOrNot Then
        'ThisWorkbook.Worksheets(g_CONTROL).Cells(g_CONTROL_DATA_L + numcb - 1, g_CONTROL_DATA_C).Font.Size = 12
    'Else
        'ThisWorkbook.Worksheets(g_CONTROL).Cells(g_CONTROL_DATA_L + numcb - 1, g_CONTROL_DATA_C).Font.Size = 10
    'End If
End Sub
Sub majListInputSheet()
    Dim ck As Integer
    ck = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.Visible = False
    
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.VALUE = False
    For k = 1 To 99
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_INPUT_L - 1, g_CONTROL_INPUT_C).VALUE = ""
    Next
    For k = 1 To selInputSheet.lvSelInputSheet.ListItems.Count
        If selInputSheet.lvSelInputSheet.ListItems(k).Checked Then
            ck = ck + 1
            ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_INPUT_L - 1, g_CONTROL_INPUT_C).VALUE = selInputSheet.lvSelInputSheet.ListItems(k).text
If ck = 1 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.Visible = True
If ck = 2 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.Visible = True
If ck = 3 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.Visible = True
If ck = 4 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.Visible = True
If ck = 5 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.Visible = True
If ck = 6 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.Visible = True
If ck = 7 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.Visible = True
If ck = 8 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.Visible = True
If ck = 9 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.Visible = True
If ck = 10 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.Visible = True
If ck = 11 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.Visible = True
If ck = 12 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.Visible = True
If ck = 13 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.Visible = True
If ck = 14 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.Visible = True
If ck = 15 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.Visible = True
If ck = 16 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.Visible = True
If ck = 17 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.Visible = True
If ck = 18 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.Visible = True
If ck = 19 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.Visible = True
If ck = 20 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.Visible = True
If ck = 21 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.Visible = True
If ck = 22 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.Visible = True
If ck = 23 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.Visible = True
If ck = 24 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.Visible = True
If ck = 25 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.Visible = True
If ck = 26 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.Visible = True
If ck = 27 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.Visible = True
If ck = 28 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.Visible = True
If ck = 29 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.Visible = True
If ck = 30 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.Visible = True
If ck = 31 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.Visible = True
If ck = 32 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.Visible = True
If ck = 33 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.Visible = True
If ck = 34 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.Visible = True
If ck = 35 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.Visible = True
If ck = 36 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.Visible = True
If ck = 37 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.Visible = True
If ck = 38 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.Visible = True
If ck = 39 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.Visible = True
If ck = 40 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.Visible = True
If ck = 41 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.Visible = True
If ck = 42 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.Visible = True
If ck = 43 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.Visible = True
If ck = 44 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.Visible = True
If ck = 45 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.Visible = True
If ck = 46 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.Visible = True
If ck = 47 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.Visible = True
If ck = 48 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.Visible = True
If ck = 49 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.Visible = True
If ck = 50 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.Visible = True
If ck = 51 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.Visible = True
If ck = 52 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.Visible = True
If ck = 53 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.Visible = True
If ck = 54 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.Visible = True
If ck = 55 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.Visible = True
If ck = 56 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.Visible = True
If ck = 57 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.Visible = True
If ck = 58 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.Visible = True
If ck = 59 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.Visible = True
If ck = 60 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.Visible = True
If ck = 61 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.Visible = True
If ck = 62 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.Visible = True
If ck = 63 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.Visible = True
If ck = 64 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.Visible = True
If ck = 65 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.Visible = True
If ck = 66 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.Visible = True
If ck = 67 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.Visible = True
If ck = 68 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.Visible = True
If ck = 69 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.Visible = True
If ck = 70 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.Visible = True
If ck = 71 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.Visible = True
If ck = 72 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.Visible = True
If ck = 73 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.Visible = True
If ck = 74 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.Visible = True
If ck = 75 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.Visible = True
If ck = 76 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.Visible = True
If ck = 77 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.Visible = True
If ck = 78 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.Visible = True
If ck = 79 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.Visible = True
If ck = 80 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.Visible = True
If ck = 81 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.Visible = True
If ck = 82 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.Visible = True
If ck = 83 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.Visible = True
If ck = 84 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.Visible = True
If ck = 85 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.Visible = True
If ck = 86 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.Visible = True
If ck = 87 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.Visible = True
If ck = 88 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.Visible = True
If ck = 89 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.Visible = True
If ck = 90 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.Visible = True
If ck = 91 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.Visible = True
If ck = 92 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.Visible = True
If ck = 93 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.Visible = True
If ck = 94 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.Visible = True
If ck = 95 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.Visible = True
If ck = 96 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.Visible = True
If ck = 97 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.Visible = True
If ck = 98 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.Visible = True
If ck = 99 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.Visible = True
        End If
    Next
End Sub
Sub initListInputSheet(list() As String)
    Dim ck As Integer
    ck = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.Visible = False
    
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.VALUE = False
    For k = 1 To 99
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_INPUT_L - 1, g_CONTROL_INPUT_C).VALUE = ""
    Next
    For k = 1 To UBound(list)
        ck = ck + 1
        ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_INPUT_L - 1, g_CONTROL_INPUT_C).VALUE = list(ck)
If ck = 1 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.Visible = True
If ck = 2 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.Visible = True
If ck = 3 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.Visible = True
If ck = 4 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.Visible = True
If ck = 5 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.Visible = True
If ck = 6 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.Visible = True
If ck = 7 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.Visible = True
If ck = 8 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.Visible = True
If ck = 9 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.Visible = True
If ck = 10 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.Visible = True
If ck = 11 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.Visible = True
If ck = 12 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.Visible = True
If ck = 13 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.Visible = True
If ck = 14 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.Visible = True
If ck = 15 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.Visible = True
If ck = 16 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.Visible = True
If ck = 17 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.Visible = True
If ck = 18 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.Visible = True
If ck = 19 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.Visible = True
If ck = 20 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.Visible = True
If ck = 21 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.Visible = True
If ck = 22 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.Visible = True
If ck = 23 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.Visible = True
If ck = 24 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.Visible = True
If ck = 25 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.Visible = True
If ck = 26 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.Visible = True
If ck = 27 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.Visible = True
If ck = 28 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.Visible = True
If ck = 29 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.Visible = True
If ck = 30 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.Visible = True
If ck = 31 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.Visible = True
If ck = 32 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.Visible = True
If ck = 33 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.Visible = True
If ck = 34 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.Visible = True
If ck = 35 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.Visible = True
If ck = 36 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.Visible = True
If ck = 37 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.Visible = True
If ck = 38 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.Visible = True
If ck = 39 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.Visible = True
If ck = 40 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.Visible = True
If ck = 41 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.Visible = True
If ck = 42 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.Visible = True
If ck = 43 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.Visible = True
If ck = 44 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.Visible = True
If ck = 45 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.Visible = True
If ck = 46 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.Visible = True
If ck = 47 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.Visible = True
If ck = 48 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.Visible = True
If ck = 49 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.Visible = True
If ck = 50 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.Visible = True
If ck = 51 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.Visible = True
If ck = 52 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.Visible = True
If ck = 53 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.Visible = True
If ck = 54 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.Visible = True
If ck = 55 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.Visible = True
If ck = 56 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.Visible = True
If ck = 57 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.Visible = True
If ck = 58 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.Visible = True
If ck = 59 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.Visible = True
If ck = 60 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.Visible = True
If ck = 61 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.Visible = True
If ck = 62 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.Visible = True
If ck = 63 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.Visible = True
If ck = 64 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.Visible = True
If ck = 65 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.Visible = True
If ck = 66 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.Visible = True
If ck = 67 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.Visible = True
If ck = 68 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.Visible = True
If ck = 69 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.Visible = True
If ck = 70 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.Visible = True
If ck = 71 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.Visible = True
If ck = 72 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.Visible = True
If ck = 73 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.Visible = True
If ck = 74 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.Visible = True
If ck = 75 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.Visible = True
If ck = 76 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.Visible = True
If ck = 77 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.Visible = True
If ck = 78 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.Visible = True
If ck = 79 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.Visible = True
If ck = 80 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.Visible = True
If ck = 81 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.Visible = True
If ck = 82 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.Visible = True
If ck = 83 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.Visible = True
If ck = 84 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.Visible = True
If ck = 85 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.Visible = True
If ck = 86 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.Visible = True
If ck = 87 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.Visible = True
If ck = 88 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.Visible = True
If ck = 89 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.Visible = True
If ck = 90 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.Visible = True
If ck = 91 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.Visible = True
If ck = 92 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.Visible = True
If ck = 93 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.Visible = True
If ck = 94 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.Visible = True
If ck = 95 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.Visible = True
If ck = 96 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.Visible = True
If ck = 97 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.Visible = True
If ck = 98 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.Visible = True
If ck = 99 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.Visible = True
    Next
End Sub
Sub majListDataSheet()
    Dim ck As Integer
    Dim tabCR() As Variant
    ck = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.Visible = False
    
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.VALUE = False
    'tabCR = ThisWorkbook.Worksheets(g_CONTROL).Range("A1:" & DecAlph(PosFct) & LigCal(1)).
    tabCR = ActiveWorkbook.Worksheets(g_CONTROL).Range(DecAlph(g_CONTROL_DATA_C) & g_CONTROL_DATA_L & ":" & DecAlph(g_CONTROL_DATA_C + 19) & (g_CONTROL_DATA_L + 19)).VALUE
    For k = 1 To 99
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C).VALUE = ""
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C + 1).VALUE = ""
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C + 2).VALUE = ""
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C + 3).VALUE = ""
    Next
    For k = 1 To selDataSheet.lvSelInputSheet.ListItems.Count
        If selDataSheet.lvSelInputSheet.ListItems(k).Checked Then
            ck = ck + 1
ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C).VALUE = selDataSheet.lvSelInputSheet.ListItems(k).text
            For Nolig = LBound(tabCR, 1) To UBound(tabCR, 1)
                If Trim(tabCR(Nolig, 1)) = selDataSheet.lvSelInputSheet.ListItems(k).text Then
                    For c = 1 To 39
ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C + c).VALUE = tabCR(Nolig, c + 1)
                    Next
                End If
            Next

If ck = 1 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.Visible = True
If ck = 2 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.Visible = True
If ck = 3 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.Visible = True
If ck = 4 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.Visible = True
If ck = 5 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.Visible = True
If ck = 6 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.Visible = True
If ck = 7 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.Visible = True
If ck = 8 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.Visible = True
If ck = 9 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.Visible = True
If ck = 10 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.Visible = True
If ck = 11 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.Visible = True
If ck = 12 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.Visible = True
If ck = 13 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.Visible = True
If ck = 14 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.Visible = True
If ck = 15 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.Visible = True
If ck = 16 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.Visible = True
If ck = 17 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.Visible = True
If ck = 18 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.Visible = True
If ck = 19 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.Visible = True
If ck = 20 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.Visible = True
If ck = 21 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.Visible = True
If ck = 22 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.Visible = True
If ck = 23 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.Visible = True
If ck = 24 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.Visible = True
If ck = 25 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.Visible = True
If ck = 26 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.Visible = True
If ck = 27 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.Visible = True
If ck = 28 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.Visible = True
If ck = 29 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.Visible = True
If ck = 30 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.Visible = True
If ck = 31 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.Visible = True
If ck = 32 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.Visible = True
If ck = 33 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.Visible = True
If ck = 34 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.Visible = True
If ck = 35 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.Visible = True
If ck = 36 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.Visible = True
If ck = 37 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.Visible = True
If ck = 38 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.Visible = True
If ck = 39 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.Visible = True
If ck = 40 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.Visible = True
If ck = 41 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.Visible = True
If ck = 42 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.Visible = True
If ck = 43 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.Visible = True
If ck = 44 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.Visible = True
If ck = 45 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.Visible = True
If ck = 46 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.Visible = True
If ck = 47 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.Visible = True
If ck = 48 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.Visible = True
If ck = 49 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.Visible = True
If ck = 50 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.Visible = True
If ck = 51 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.Visible = True
If ck = 52 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.Visible = True
If ck = 53 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.Visible = True
If ck = 54 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.Visible = True
If ck = 55 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.Visible = True
If ck = 56 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.Visible = True
If ck = 57 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.Visible = True
If ck = 58 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.Visible = True
If ck = 59 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.Visible = True
If ck = 60 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.Visible = True
If ck = 61 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.Visible = True
If ck = 62 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.Visible = True
If ck = 63 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.Visible = True
If ck = 64 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.Visible = True
If ck = 65 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.Visible = True
If ck = 66 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.Visible = True
If ck = 67 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.Visible = True
If ck = 68 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.Visible = True
If ck = 69 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.Visible = True
If ck = 70 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.Visible = True
If ck = 71 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.Visible = True
If ck = 72 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.Visible = True
If ck = 73 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.Visible = True
If ck = 74 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.Visible = True
If ck = 75 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.Visible = True
If ck = 76 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.Visible = True
If ck = 77 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.Visible = True
If ck = 78 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.Visible = True
If ck = 79 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.Visible = True
If ck = 80 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.Visible = True
If ck = 81 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.Visible = True
If ck = 82 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.Visible = True
If ck = 83 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.Visible = True
If ck = 84 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.Visible = True
If ck = 85 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.Visible = True
If ck = 86 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.Visible = True
If ck = 87 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.Visible = True
If ck = 88 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.Visible = True
If ck = 89 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.Visible = True
If ck = 90 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.Visible = True
If ck = 91 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.Visible = True
If ck = 92 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.Visible = True
If ck = 93 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.Visible = True
If ck = 94 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.Visible = True
If ck = 95 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.Visible = True
If ck = 96 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.Visible = True
If ck = 97 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.Visible = True
If ck = 98 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.Visible = True
If ck = 99 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.Visible = True

        End If
    Next
End Sub
Sub initListMetaSheet(list() As String)
    Dim ck As Integer
    ck = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta1.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta2.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta3.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta4.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta5.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta6.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta7.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta8.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta9.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta10.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta11.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta12.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta13.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta14.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta15.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta16.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta17.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta18.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta19.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta20.Visible = False
    
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta1.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta2.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta3.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta4.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta5.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta6.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta7.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta8.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta9.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta10.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta11.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta12.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta13.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta14.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta15.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta16.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta17.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta18.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta19.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta20.VALUE = False
    For k = 1 To 20
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_META_L - 1, g_CONTROL_META_C).VALUE = ""
    Next

    For k = 1 To UBound(list)
        ck = ck + 1
        ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_META_L - 1, g_CONTROL_META_C).VALUE = list(ck)
If ck = 1 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta1.Visible = True
If ck = 2 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta2.Visible = True
If ck = 3 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta3.Visible = True
If ck = 4 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta4.Visible = True
If ck = 5 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta5.Visible = True
If ck = 6 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta6.Visible = True
If ck = 7 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta7.Visible = True
If ck = 8 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta8.Visible = True
If ck = 9 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta9.Visible = True
If ck = 10 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta10.Visible = True
If ck = 11 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta11.Visible = True
If ck = 12 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta12.Visible = True
If ck = 13 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta13.Visible = True
If ck = 14 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta14.Visible = True
If ck = 15 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta15.Visible = True
If ck = 16 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta16.Visible = True
If ck = 17 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta17.Visible = True
If ck = 18 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta18.Visible = True
If ck = 19 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta19.Visible = True
If ck = 20 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta20.Visible = True
    Next
End Sub
Sub initListDataSheet(list() As String)
    Dim ck As Integer
    ck = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.Visible = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.Visible = False
    
    
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.VALUE = False
    ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.VALUE = False
    For k = 1 To 99
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C).VALUE = ""
        ThisWorkbook.Worksheets(g_CONTROL).Cells(k + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_GEN_C).Interior.Color = RGB(244, 176, 132)
    Next
'MsgBox LBound(list) & UBound(list) & Chr(10) & Join(list, Chr(10))
    For k = 1 To UBound(list)
        ck = ck + 1
        ThisWorkbook.Worksheets(g_CONTROL).Cells(ck + g_CONTROL_DATA_L - 1, g_CONTROL_DATA_C).VALUE = list(ck)
If ck = 1 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.Visible = True
If ck = 2 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.Visible = True
If ck = 3 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.Visible = True
If ck = 4 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.Visible = True
If ck = 5 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.Visible = True
If ck = 6 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.Visible = True
If ck = 7 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.Visible = True
If ck = 8 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.Visible = True
If ck = 9 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.Visible = True
If ck = 10 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.Visible = True
If ck = 11 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.Visible = True
If ck = 12 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.Visible = True
If ck = 13 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.Visible = True
If ck = 14 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.Visible = True
If ck = 15 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.Visible = True
If ck = 16 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.Visible = True
If ck = 17 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.Visible = True
If ck = 18 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.Visible = True
If ck = 19 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.Visible = True
If ck = 20 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.Visible = True
If ck = 21 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.Visible = True
If ck = 22 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.Visible = True
If ck = 23 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.Visible = True
If ck = 24 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.Visible = True
If ck = 25 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.Visible = True
If ck = 26 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.Visible = True
If ck = 27 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.Visible = True
If ck = 28 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.Visible = True
If ck = 29 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.Visible = True
If ck = 30 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.Visible = True
If ck = 31 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.Visible = True
If ck = 32 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.Visible = True
If ck = 33 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.Visible = True
If ck = 34 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.Visible = True
If ck = 35 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.Visible = True
If ck = 36 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.Visible = True
If ck = 37 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.Visible = True
If ck = 38 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.Visible = True
If ck = 39 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.Visible = True
If ck = 40 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.Visible = True
If ck = 41 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.Visible = True
If ck = 42 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.Visible = True
If ck = 43 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.Visible = True
If ck = 44 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.Visible = True
If ck = 45 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.Visible = True
If ck = 46 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.Visible = True
If ck = 47 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.Visible = True
If ck = 48 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.Visible = True
If ck = 49 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.Visible = True
If ck = 50 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.Visible = True
If ck = 51 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.Visible = True
If ck = 52 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.Visible = True
If ck = 53 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.Visible = True
If ck = 54 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.Visible = True
If ck = 55 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.Visible = True
If ck = 56 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.Visible = True
If ck = 57 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.Visible = True
If ck = 58 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.Visible = True
If ck = 59 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.Visible = True
If ck = 60 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.Visible = True
If ck = 61 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.Visible = True
If ck = 62 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.Visible = True
If ck = 63 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.Visible = True
If ck = 64 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.Visible = True
If ck = 65 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.Visible = True
If ck = 66 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.Visible = True
If ck = 67 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.Visible = True
If ck = 68 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.Visible = True
If ck = 69 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.Visible = True
If ck = 70 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.Visible = True
If ck = 71 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.Visible = True
If ck = 72 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.Visible = True
If ck = 73 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.Visible = True
If ck = 74 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.Visible = True
If ck = 75 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.Visible = True
If ck = 76 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.Visible = True
If ck = 77 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.Visible = True
If ck = 78 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.Visible = True
If ck = 79 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.Visible = True
If ck = 80 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.Visible = True
If ck = 81 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.Visible = True
If ck = 82 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.Visible = True
If ck = 83 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.Visible = True
If ck = 84 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.Visible = True
If ck = 85 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.Visible = True
If ck = 86 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.Visible = True
If ck = 87 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.Visible = True
If ck = 88 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.Visible = True
If ck = 89 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.Visible = True
If ck = 90 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.Visible = True
If ck = 91 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.Visible = True
If ck = 92 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.Visible = True
If ck = 93 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.Visible = True
If ck = 94 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.Visible = True
If ck = 95 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.Visible = True
If ck = 96 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.Visible = True
If ck = 97 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.Visible = True
If ck = 98 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.Visible = True
If ck = 99 Then ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.Visible = True

    Next
End Sub
Sub listUpdateFromModel()
    On Error GoTo errorHandler
    Dim nbData As Integer
    Dim nbInput As Integer
    Dim bt As Integer
    Dim ba As Integer
    Dim bs As Integer
    Dim bn As Integer
    Dim bq As Integer
    
    If Not openFileIfNot("MODELE") Then GoTo errorHandler
    nbData = 0
    nbInput = 0
    bt = ba = bs = bn = bq = 0
    ThisWorkbook.Worksheets(g_CONTROL).cbTime.Clear
    ThisWorkbook.Worksheets(g_CONTROL).cbArea.Clear
    ThisWorkbook.Worksheets(g_CONTROL).cbScenario.Clear
    ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.Clear
    ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.Clear
    'ActiveSheet.ListData.Clear
    'ActiveSheet.ListInput.Clear
    For Each WS In g_WB_Modele.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        If firstcell Like "*DATE*" Then
            bt = bt + 1
            ThisWorkbook.Worksheets(g_CONTROL).cbTime.AddItem (WS.NAME)
        End If
        If firstcell Like "*AREA*" Then
            ba = ba + 1
            ThisWorkbook.Worksheets(g_CONTROL).cbArea.AddItem (WS.NAME)
        End If
        If firstcell Like "*SCENARIO*" Then
            bs = bs + 1
            ThisWorkbook.Worksheets(g_CONTROL).cbScenario.AddItem (WS.NAME)
        End If
        If firstcell Like "*NOMENCLATURE*" Then
            bn = bn + 1
            ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.AddItem (WS.NAME)
        End If
        If firstcell Like "*QUANTITE*" Then
            bq = bq + 1
            ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.AddItem (WS.NAME)
        End If
        'If firstcell Like "*NOP_Col*" Then
            'nbData = nbData + 1
        'End If
        'If firstcell Like "*INPUT*" Then
            'nbInput = nbInput + 1
        'End If
    Next WS
    If bt > 0 Then
        ThisWorkbook.Worksheets(g_CONTROL).cbTime.ListIndex = 0
    End If
    If ba > 0 Then
        ThisWorkbook.Worksheets(g_CONTROL).cbArea.ListIndex = 0
    End If
    If bs > 0 Then
        ThisWorkbook.Worksheets(g_CONTROL).cbScenario.ListIndex = 0
    End If
    If bn > 0 Then
        ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.ListIndex = 0
        If WsExist(g_NOMENCLATURE, g_WB_Modele) Then
            Set FLQ = g_WB_Modele.Sheets(g_NOMENCLATURE)
            dl = getDerLig(FLQ)
            If dl > 1 Then
                For I = 0 To ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.ListCount - 1
                    If ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.list(I) = FLQ.Cells(2, 1).VALUE Then
                        ThisWorkbook.Worksheets(g_CONTROL).cbNomenclature.ListIndex = I
                    End If
                Next
            End If
        End If
    End If
    If bq > 0 Then
        ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.ListIndex = 0
        If WsExist(g_QUANTITY, g_WB_Modele) Then
            Set FLQ = g_WB_Modele.Sheets(g_QUANTITY)
            dl = getDerLig(FLQ)
            If dl > 1 Then
                For I = 0 To ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.ListCount - 1
                    If ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.list(I) = FLQ.Cells(2, 1).VALUE Then
                        ThisWorkbook.Worksheets(g_CONTROL).cbQuantites.ListIndex = I
                    End If
                Next
            End If
        End If
    End If
    'ThisWorkbook.Worksheets(g_CONTROL).butSelInputSheet.Caption = "Sélection / " & nbInput
    'ThisWorkbook.Worksheets(g_CONTROL).butSelDataSheet.Caption = "Sélection / " & nbData
    Call initInfo
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "listUpdateFromModel"): Exit Sub
End Sub
Sub listUpdateFromModelSheet()
    On Error GoTo errorHandler
    Dim nbMeta As Integer
    Dim nbData As Integer
    Dim nbInput As Integer
    Dim Mnames() As String
    Dim Dnames() As String
    Dim Inames() As String
    Dim firstcell As String
    If Not openFileIfNot("MODELE") Then GoTo errorHandler
    ReDim Mnames(0)
    ReDim Dnames(0)
    ReDim Inames(0)
    nbData = 0
    nbMeta = 0
    nbInput = 0
    'ActiveSheet.ListData.Clear
    'ActiveSheet.ListInput.Clear
    For Each WS In g_WB_Modele.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        If firstcell Like "*META*" Then
            nbMeta = nbMeta + 1
            Mnames = addToList(Mnames, WS.NAME)
        End If
        If firstcell Like "*NOP_Col*" Then
            nbData = nbData + 1
            Dnames = addToList(Dnames, WS.NAME)
        End If
        If firstcell Like "*INPUT*" Then
            nbInput = nbInput + 1
            Inames = addToList(Inames, WS.NAME)
        End If
    Next WS
    g_nbInput = nbInput
    g_nbMeta = nbMeta
    g_nbData = nbData
    ThisWorkbook.Worksheets(g_CONTROL).butSelInputSheet.Caption = "Sélection / " & nbInput
    ThisWorkbook.Worksheets(g_CONTROL).butSelMetaSheet.Caption = "Sélection / " & nbMeta
    ThisWorkbook.Worksheets(g_CONTROL).butSelDataSheet.Caption = "Sélection / " & nbData
    Call initListMetaSheet(Mnames)
    Call initListDataSheet(Dnames)
    Call initListInputSheet(Inames)
    Call initInfo
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "listUpdateFromModelSheet"): Exit Sub
End Sub
Sub listUpdateGenFromModelSheet(listD() As String)
    On Error GoTo errorHandler
    Dim nbData As Integer
    Dim Dnames() As String
    Dim firstcell As String
    If Not openFileIfNot("MODELE") Then GoTo errorHandler
    ReDim Dnames(0)
    nbData = 0
    For Each WS In g_WB_Modele.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        If firstcell Like "*NOP_Col*" Then
            nbData = nbData + 1
            Dnames = addToList(Dnames, WS.NAME)
        End If
    Next WS
    ThisWorkbook.Worksheets(g_CONTROL).butSelDataSheet.Caption = "Sélection / " & nbData
    Call initListDataSheet(Dnames)
    Call initGenDatInfo
    Call colorGenDat(listD)
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "listUpdateGenFromModelSheet"): Exit Sub
End Sub
Sub colorGenDat(listD() As String)
    Dim cellValue As String
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set FLCTL = g_WB_Extra.Worksheets(g_CONTROL)
    For I = 0 To 99
        cellValue = FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C - 1).VALUE
        If stringIsInList(cellValue, listD) Then
            FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C).Interior.Color = RGB(169, 208, 142)
        Else
            FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C).Interior.Color = RGB(244, 176, 132)
        End If
    Next
End Sub
Sub listUpdate()
    Call DisableExcel
    Call listUpdateFromModel
    Call listUpdateFromModelSheet
    Call EnableExcel
End Sub
Sub initInfo()
    Dim FLCTL As Worksheet
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set FLCTL = g_WB_Extra.Worksheets(g_CONTROL)
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C).VALUE = ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 1).VALUE = ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 2).VALUE = ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 3).VALUE = ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C).VALUE = ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 1).VALUE = ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 2).VALUE = ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 3).VALUE = ""
    For I = 0 To 99
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_INPUT_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_INPUT_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_INPUT_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_INPUT_C + 3).VALUE = ""
        'FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_META_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_META_L + I, g_CONTROL_META_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_META_L + I, g_CONTROL_META_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_META_L + I, g_CONTROL_META_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_META_L + I, g_CONTROL_META_C + 4).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 3).VALUE = ""
    Next
End Sub
Sub initGenDatInfo()
    Dim FLCTL As Worksheet
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set FLCTL = g_WB_Extra.Worksheets(g_CONTROL)
    For I = 0 To 99
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_GEN_L + I, g_CONTROL_DATA_GEN_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_DAT_L + I, g_CONTROL_DATA_DAT_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_FOR_L + I, g_CONTROL_DATA_FOR_C + 3).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 1).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 2).VALUE = ""
        FLCTL.Cells(g_CONTROL_DATA_COM_L + I, g_CONTROL_DATA_COM_C + 3).VALUE = ""
    Next
End Sub


