Attribute VB_Name = "File"
Option Explicit
''' Sur click du bouton cancel de PARAMETERS, SOURCE ou CIBLE 011016
Public Sub cancelFileTyp(ByVal fichier As String)
    If fichier = "PARAMETERS" Then
        ThisWorkbook.Worksheets(g_CONTROL).labParametersFile.Caption = "Le modèle est dans ce fichier"
        ThisWorkbook.Worksheets(g_CONTROL).labParametersFileName.Caption = ThisWorkbook.NAME
        Set g_WB_Modele = ThisWorkbook
        g_WB_Modele_Name = ThisWorkbook.NAME
        Call listUpdateFromModel
    End If
    If fichier = "SOURCE" Then
        ThisWorkbook.Worksheets(g_CONTROL).labSourceFile.Caption = "Les données source sont dans ce fichier"
        ThisWorkbook.Worksheets(g_CONTROL).labSourceFileName.Caption = ThisWorkbook.NAME
        Set g_WB_Source = ThisWorkbook
        g_WB_Source_Name = ThisWorkbook.NAME
    End If
    If fichier = "TARGET" Then
        ThisWorkbook.Worksheets(g_CONTROL).labTargetFile.Caption = "Les feuilles cible sont dans ce fichier"
        ThisWorkbook.Worksheets(g_CONTROL).labTargetFileName.Caption = ThisWorkbook.NAME
        Set g_WB_Target = ThisWorkbook
        g_WB_Target_Name = ThisWorkbook.NAME
        Call listUpdateFromModelSheet
    End If
End Sub
''' Sur click du bouton PARAMETERS, SOURCE ou CIBLE 011016
Public Sub openfileTypFile(ByVal fichier As String)
    Dim oFD As Object
    Dim FilePath As String
    Dim FileBefore As String
    Dim FileName As String
    Dim libelle As String
    Dim b As Boolean
    
    If fichier = "PARAMETERS" Then libelle = "des paramètres"
    If fichier = "SOURCE" Then libelle = "source"
    If fichier = "TARGET" Then libelle = "cible"
    Set oFD = Application.FileDialog(msoFileDialogOpen)
    oFD.Title = "Selectionner le fichier " & libelle
    oFD.AllowMultiSelect = False ''Disable multiSelection
    ''Apply filter on Excel
    oFD.Filters.Clear
    oFD.Filters.Add "Excel file", "*.xls; *.xlsx; *.xlsm"
    oFD.FilterIndex = 1
    If (oFD.Show()) Then
        FilePath = oFD.SelectedItems(1)
        FileName = Split(FilePath, "\")(UBound(Split(FilePath, "\")))
        If fichier = "PARAMETERS" Then g_PARAMETERS = FilePath
        If fichier = "SOURCE" Then g_SOURCE = FilePath
        If fichier = "TARGET" Then g_TARGET = FilePath
        FileBefore = Left(FilePath, Len(FilePath) - Len(FileName))
        If fichier = "PARAMETERS" Then
            ThisWorkbook.Worksheets(g_CONTROL).labParametersFile.Caption = FileBefore
            ThisWorkbook.Worksheets(g_CONTROL).labParametersFileName.Caption = FileName
            g_WB_Modele_Path = FilePath
            g_WB_Modele_Name = FileName
            b = openFileIfNot("MODELE")
            '''Call listUpdateFromModel
            '''Call listUpdateFromModelSheet
        End If
        If fichier = "SOURCE" Then
            ThisWorkbook.Worksheets(g_CONTROL).labSourceFile.Caption = FileBefore
            ThisWorkbook.Worksheets(g_CONTROL).labSourceFileName.Caption = FileName
            g_WB_Source_Path = FilePath
            g_WB_Source_Name = FileName
            b = openFileIfNot("SOURCE")
            ' maj des listes (colorer la ligne où unINPUT est possible)?
        End If
        If fichier = "TARGET" Then
            ThisWorkbook.Worksheets(g_CONTROL).labTargetFile.Caption = FileBefore
            ThisWorkbook.Worksheets(g_CONTROL).labTargetFileName.Caption = FileName
            g_WB_Target_Path = FilePath
            g_WB_Target_Name = FileName
            b = openFileIfNot("TARGET")
            '''Call listUpdateFromTarget
        End If
    End If
    Set oFD = Nothing
End Sub
''' Vérification si un fichier est ouvert et ouverture sinon
''' False si le fichier n'existe pas ou plus
Public Function openFileIfNot(nom As String) As Boolean
    If nom = "MODELE" Then
        If ThisWorkbook.NAME = g_WB_Modele_Name Then
            ' Le modèle est dans le fichier Extra
            openFileIfNot = True
            Set g_WB_Modele = ThisWorkbook
        Else
            g_WB_Modele_Name = ThisWorkbook.Worksheets(g_CONTROL).labParametersFileName.Caption
            g_WB_Modele_Path = ThisWorkbook.Worksheets(g_CONTROL).labParametersFile.Caption _
            & ThisWorkbook.Worksheets(g_CONTROL).labParametersFileName.Caption
            If testFileOpen(g_WB_Modele_Name) Then
                ' Fichier déjà ouvert
                openFileIfNot = True
                Set g_WB_Modele = Workbooks(g_WB_Modele_Name)
            Else
                If FileThere(g_WB_Modele_Path) Then
                    Application.Visible = False
                    Set g_WB_Modele = Excel.Application.Workbooks.Open(g_WB_Modele_Path)
                    Application.WindowState = xlMinimized
                    Application.Visible = True
                    g_WB_Extra.Worksheets(g_CONTROL).Activate
                    openFileIfNot = True
                Else
                    MsgBox "Le fichier : " & g_WB_Modele_Path & " n'existe plus."
                    openFileIfNot = False
                End If
            End If
        End If
    End If
    If nom = "SOURCE" Then
        If ThisWorkbook.NAME = g_WB_Target_Name Then
            ' La source est dans le fichier Extra
            openFileIfNot = True
            Set g_WB_Source = ThisWorkbook
        Else
            g_WB_Source_Name = ThisWorkbook.Worksheets(g_CONTROL).labSourceFileName.Caption
            g_WB_Source_Path = ThisWorkbook.Worksheets(g_CONTROL).labSourceFile.Caption _
            & ThisWorkbook.Worksheets(g_CONTROL).labSourceFileName.Caption
            If testFileOpen(g_WB_Source_Name) Then
                ' Fichier déjà ouvert
                openFileIfNot = True
                Set g_WB_Source = Workbooks(g_WB_Source_Name)
'MsgBox "source deja ouvert=" & g_WB_Source.NAME
            Else
                If FileThere(g_WB_Source_Path) Then
                    Application.Visible = False
                    Set g_WB_Source = Excel.Application.Workbooks.Open(g_WB_Source_Path)
                    Application.Visible = True
                    g_WB_Extra.Worksheets(g_CONTROL).Activate
                    openFileIfNot = True
                Else
                    MsgBox "Le fichier : " & g_WB_Source_Path & " n'existe plus."
                    openFileIfNot = False
                End If
            End If
        End If
    End If
    If nom = "TARGET" Then
        If ThisWorkbook.NAME = g_WB_Target_Name Then
            ' La cible est dans le fichier Extra
            openFileIfNot = True
            Set g_WB_Target = ThisWorkbook
        Else
            g_WB_Target_Name = ThisWorkbook.Worksheets(g_CONTROL).labTargetFileName.Caption
            g_WB_Target_Path = ThisWorkbook.Worksheets(g_CONTROL).labTargetFile.Caption _
            & ThisWorkbook.Worksheets(g_CONTROL).labTargetFileName.Caption
            If testFileOpen(g_WB_Target_Name) Then
                ' Fichier déjà ouvert
                openFileIfNot = True
                Set g_WB_Target = Workbooks(g_WB_Target_Name)
            Else
                If FileThere(g_WB_Target_Path) Then
                    Application.Visible = False
                    Set g_WB_Target = Excel.Application.Workbooks.Open(g_WB_Target_Path)
                    Application.Visible = True
                    g_WB_Extra.Worksheets(g_CONTROL).Activate
                    openFileIfNot = True
                Else
                    MsgBox "Le fichier : " & g_WB_Target_Path & " n'existe plus."
                    openFileIfNot = False
                End If
            End If
        End If
    End If
'MsgBox nom & ":" & openFileIfNot
End Function
Function FileThere(FileName As String) As Boolean
     FileThere = (Dir(FileName) > "")
End Function
Private Function testFileOpen(nom As String) As Boolean
    Dim Classeur As Workbook
    For Each Classeur In Application.Workbooks
        If Classeur.NAME = nom Then
            testFileOpen = True
            Exit Function
        End If
    Next Classeur
    testFileOpen = False
End Function
'' On button clic CIBLE
Public Sub openfileCibleFile()
    Dim oFD As Object
    
    Set oFD = Application.FileDialog(msoFileDialogOpen)
    oFD.Title = "Selectionner le fichier cible"
    oFD.AllowMultiSelect = False ''Disable multiSelection
    ''Apply filter on Excel
    oFD.Filters.Clear
    oFD.Filters.Add "Excel file", "*.xls; *.xlsx; *.xlsm"
    oFD.FilterIndex = 1
    If (oFD.Show()) Then ThisWorkbook.Worksheets(g_CONTROL).lbl_pathCible.Caption = oFD.SelectedItems(1) Else ThisWorkbook.Worksheets(g_CONTROL).lbl_pathCible.Caption = g_noFileTxt
    Set oFD = Nothing
End Sub


'' On button clic "Rafraichir typage Onglet..."
Public Sub RefreshIHM()
    '' Clean the name sheet spaces
    Call DeletePreviousSheet
    '' Add the sheets to the godd space in function of their identity
    Call DetectActualSheet
    '' Compare with the Cible and Source sheets the current sheets
    Call CompareFiles
End Sub


Private Sub DeletePreviousSheet()

    Dim I As Long
    '' for each sheets spaces delete all
    I = g_line_nom
    With ThisWorkbook.Worksheets(g_CONTROL)
        Do While StrComp(.Cells(I, g_col_nom), vbNullString) <> 0
            .Cells(I, g_col_nom).VALUE = vbNullString
            .Cells(I, g_col_nom + 1).VALUE = vbNullString
            I = I + 1
        Loop
        I = g_line_qua
        Do While StrComp(.Cells(I, g_Col_qua), vbNullString) <> 0
            .Cells(I, g_Col_qua).VALUE = vbNullString
            .Cells(I, g_Col_qua + 1).VALUE = vbNullString
            I = I + 1
        Loop
        I = g_line_nontype
        Do While StrComp(.Cells(I, g_col_nontype), vbNullString) <> 0
            .Cells(I, g_col_nontype).VALUE = vbNullString
            .Cells(I, g_col_nontype + 1).VALUE = vbNullString
            I = I + 1
        Loop
        I = g_line_feu
        Do While StrComp(.Cells(I, g_col_feu), vbNullString) <> 0
            .Cells(I, g_col_feu).VALUE = vbNullString
            .Cells(I, g_col_feu + 1).VALUE = vbNullString
            .Cells(I, g_col_feu + 2).VALUE = vbNullString
            .Cells(I, g_col_feu + 3).VALUE = vbNullString
            .Cells(I, g_col_feu + 4).VALUE = vbNullString
            I = I + 1
        Loop
    End With
End Sub


Private Sub CompareFiles()

    Dim pathCible As String, pathSource As String
    '' Check if the path is correct for both workbooks
    With ThisWorkbook.Worksheets(g_CONTROL)
        If (True = fileExists(.lbl_pathSource.Caption)) Then
            pathCible = .lbl_pathSource.Caption
            Call FindCorrespondance(pathCible, g_SOURCE)
        End If
        If (True = fileExists(.lbl_pathCible.Caption)) Then
            pathSource = .lbl_pathCible.Caption
            Call FindCorrespondance(pathSource, g_CIBLE)
        End If
    End With
End Sub


Public Function getPath(ByVal TypePath As String) As String
    With ThisWorkbook.Worksheets(g_CONTROL)
        If (0 = StrComp(TypePath, g_CIBLE)) Then
            getPath = .lbl_pathCible.Caption
        ElseIf (0 = StrComp(TypePath, g_SOURCE)) Then
            getPath = .lbl_pathSource.Caption
        Else
            getPath = vbNullString
        End If
    End With
End Function


Private Sub FindCorrespondance(ByVal pathFile As String, ByVal Filetype As Long)
    
    Dim xlApp As New Excel.Application
    Dim Target_WB As Workbook
    Dim I As Long, j As Long, colNumber As Long
    Dim found As Boolean
    
    '' Compare the name of the different sheets of the current workBook and the target workbook ( Cible or Source )
    Set Target_WB = xlApp.Workbooks.Open(pathFile)
    If (g_CIBLE = Filetype) Then
        colNumber = g_col_feu_cible
    ElseIf (g_SOURCE = Filetype) Then
        colNumber = g_col_feu_source
    End If
    found = False
     I = g_line_feu
    With ThisWorkbook.Worksheets(g_CONTROL)
        Do While StrComp(.Cells(I, g_col_feu), vbNullString) <> 0
            For j = 1 To Target_WB.Sheets.Count
                If (0 = StrComp(Target_WB.Sheets(j).NAME, .Cells(I, g_col_feu))) Then
                    found = True
                    Exit For
                End If
            Next j
            If (found = True) Then .Cells(I, colNumber).VALUE = g_Sheet_Found Else .Cells(I, colNumber).VALUE = g_Sheet_NotFound
            I = I + 1
            found = False
        Loop
    End With
    Target_WB.Close (False)
    xlApp.Quit
    Set Target_WB = Nothing
    Set xlApp = Nothing
    
End Sub


Private Sub ResizeIHMTab(ByVal nbLine As Long)
    ThisWorkbook.Worksheets(g_CONTROL).ListObjects(1).Resize Range(Cells(g_line_feu - 1, g_col_feu), Cells(g_line_feu + nbLine - 1, g_col_feu + 4))
End Sub

'' NOMENCLATURE = NOMENCLATURE sheet
'' QUANTITE = QUANTITE Sheet
'' DATA = feuille sheet
Private Sub DetectActualSheet()
    Dim nbNomenclature As Long, nbQuantite As Long, nbFeuille As Long, nbNontype As Long, I As Long
    
    nbNomenclature = 0
    nbQuantite = 0
    nbFeuille = 0
    nbNontype = 0
    
    '' For each sheet, detect the different A1 cells and write in the good sheets name spaces
    For I = 1 To ThisWorkbook.Sheets.Count
        With ThisWorkbook.Sheets(I)
            If (0 = StrComp(.Cells(1, 1), g_NOMENCLATURE)) Then
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_nom + nbNomenclature, g_col_nom) = .NAME
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_nom + nbNomenclature, g_col_nom + 2) = "OK"
                nbNomenclature = nbNomenclature + 1
                If (.Shapes.Count < 2) Then Call AddButtonNomenclature(.NAME)
            ElseIf (0 = StrComp(.Cells(1, 1), g_QUANTITE)) Then
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_qua + nbQuantite, g_Col_qua) = .NAME
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_qua + nbQuantite, g_Col_qua + 2) = "OK"
                nbQuantite = nbQuantite + 1
                If (.Shapes.Count < 2) Then Call AddButtonQuantite(.NAME)
            ElseIf (0 = StrComp(.Cells(1, 1), g_feuille)) Then
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_feu + nbFeuille, g_col_feu) = .NAME
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_feu + nbFeuille, g_col_feu + 2) = "OK"
                nbFeuille = nbFeuille + 1
            Else
                ThisWorkbook.Sheets(g_CONTROL).Cells(g_line_nontype + nbNontype, g_col_nontype) = .NAME
                nbNontype = nbNontype + 1
            End If
        End With
    Next I
    '' Resize the table of the data sheet
    Call ResizeIHMTab(nbFeuille)
End Sub


Private Sub AddButtonNomenclature(ByVal SheetName As String)
    With ThisWorkbook.Sheets(SheetName)
        .Rows(g_spaceButton_line & ":" & g_spaceButton_line).RowHeight = 30
        .Buttons.Add(10, 5, 130, 35).Select
        .Buttons(1).OnAction = "NomenclatureIsolee"
        .Buttons(1).Caption = "Génerer Nomenclature Isolee"
        .Buttons.Add(150, 5, 130, 35).Select
        .Buttons(2).OnAction = "NomenclatureIncrmentale"
        .Buttons(2).Caption = "Génerer Nomenclature Incrementale"
    End With
End Sub

Private Sub AddButtonQuantite(ByVal SheetName As String)
    With ThisWorkbook.Sheets(SheetName)
        .Rows(g_spaceButton_line & ":" & g_spaceButton_line).RowHeight = 30
        .Buttons.Add(10, 5, 130, 35).Select
        .Buttons(1).OnAction = "QuantiteIsolee"
        .Buttons(1).Caption = "Génerer Quantité Isolee"
        .Buttons.Add(150, 5, 130, 35).Select
        .Buttons(2).OnAction = "QuantiteIncrmentale"
        .Buttons(2).Caption = "Génerer Quantité Incrementale"
    End With
End Sub

Public Sub DisableExcelFunction()
    'Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub

Public Sub EnableExcelFunction()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.Cursor = xlDefault
End Sub

