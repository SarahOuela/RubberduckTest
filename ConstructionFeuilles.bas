Attribute VB_Name = "ConstructionFeuilles"
' Variables globales
Dim logNom() As Variant
Dim ATTRIBUTS() As String
Dim NUMENCOURS As Integer
Dim testCode As Boolean
Dim ACTION As Integer, ID As Integer, BEFORE1 As Integer, THIS1 As Integer, ATTR1 As Integer, _
     BEFORE2 As Integer, THIS2 As Integer, ATTR2 As Integer, _
     BEFORE3 As Integer, THIS3 As Integer, ATTR3 As Integer, _
     BEFORE4 As Integer, THIS4 As Integer, ATTR4 As Integer, _
     BEFORE5 As Integer, THIS5 As Integer, ATTR5 As Integer, _
     BEFORE6 As Integer, THIS6 As Integer, ATTR6 As Integer, _
     BEFORE7 As Integer, THIS7 As Integer, ATTR7 As Integer, _
     BEFORE8 As Integer, THIS8 As Integer, ATTR8 As Integer, _
     BEFORE9 As Integer, THIS9 As Integer, ATTR9 As Integer, _
     EXTENDED As Integer, NAME As Integer, quantity As Integer, _
     AREA As Integer, temps As Integer, equation As Integer, _
     SCALE0 As Integer, UNIT As Integer, SUBSTITUTE As Integer, _
     DEFAUT As Integer, VALUE As Integer, BOUCLE As Integer, _
     SCENARIO As Integer
'
Dim testPileRecursivite As String
' Liste des fonctions autorisées utile notamment lors de l'analyse syntaxique
Dim functionsList() As String
'
Dim FLCONTROL As Worksheet
Dim FLTIM As Worksheet
Dim FL0NO As Worksheet
Dim FLNOM As Worksheet
Dim FL0EQ As Worksheet
Dim FLEQU As Worksheet
Dim FLSSK As Worksheet
Dim FLSAI As Worksheet
Dim FLCSK As Worksheet
Dim FLCAL As Worksheet
Dim FLOLD As Worksheet
Dim FLOHM As Worksheet
Dim FLHMA As Worksheet
Dim FL0HM As Worksheet
'
Dim TIMEARRAY() As Single
Dim FORMULESLOCALES() As Variant
Dim QUANTITECAL() As Variant
Dim NOMENCLATURE() As String    ' La nomenclature étendue
Dim EQUATIONS() As String       ' Les périmètres des équations étendues
Dim EQUATIONSQ() As String      ' Les quantités des équations étendues
Dim EQUATIONSF() As String      ' Les formules des équations étendues
'Dim CLASSES() As String         ' Uniquement les classes de la nomenclature
Dim CLASSES() As String
Dim CLASSESI() As Variant
Dim CLASSESN() As Integer
Dim LISTFCT() As String         ' Liste des fonctions réservées
Dim POURSUIVRE As Boolean       ' Pour  interrompre l'enchainement des Sub si nécessaire
Dim Etape As Integer
Dim START As Single
Dim CONTEXTE As String
Dim NOMBRE As Integer

Dim COLOK As Integer
Dim COLRESULTAT As Integer
Dim COLRESULTATNB As Integer
Dim COLINPUTNB As Integer
Dim COLTIME As Integer

Dim LigActionControl As Integer
Dim DerLigControl As Integer
Dim ColResultatControl As Integer
Function addItemToList(Item As String, list() As String) As String()
    ' Ajoute à la fin item à la liste list
    ' ne fait rien si iem est déjà dans la liste
    Dim test As Boolean
    testaAttr = False
    For la = LBound(list) To UBound(list)
        If list(la) = Item Then
            test = True
            Exit For
        End If
    Next
    If Not test Then
        ' alimenter la liste
        If UBound(list) = 0 Then
            ReDim list(1 To 1)
        Else
            ReDim Preserve list(1 To UBound(list) + 1)
        End If
        list(UBound(list)) = Item
    End If
    addItemToList = list
End Function
Function addItemToEnd(Item As String, list() As String) As String()
    ' Ajoute à la fin item à la liste list
    If UBound(list) = 0 Then
        ReDim list(1 To 1)
    Else
        ReDim Preserve list(1 To UBound(list) + 1)
    End If
    list(UBound(list)) = Item
    addItemToEnd = list
End Function
Function ecritureLog(FL As Worksheet)
    logNom = Application.Transpose(logNom)
    FL.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    ecritureLog = coloriageLog(FL)
End Function

Sub gestionModel()
    initGestionModel.Show
End Sub
Function epurer(debut As String, cellule As String) As String
    Dim clalu As String
    clalu = cellule
    If cellule Like "*.1*" Then
        If debut Like "*.1*" Then
            clalu = Replace(cellule, ".1", "")
        Else
            ' enlever les ...1 si .2 existe tout sinon
            clalu = "" 'replace(Cla(lig, cc),dd,"")
        End If
    End If
    epurer = clalu
End Function

Sub initialisation()
' Efface toutes les feuilles intermédiaires et résultats et initialise des variables globales
    Etape = 1
    On Error GoTo errorHandler
    Call SetFl
    'initialisation des feuilles de controle
    DerLigControl = Split(FLCONTROL.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLCONTROL.UsedRange.Address, "$")(3)).Column
    LigActionControl = 100
    COLOK = 3
    COLRESULTAT = 5
    COLRESULTATNB = 7
    COLINPUTNB = 10
    COLTIME = 11
    For Nolig = 1 To DerLigControl
        If FLCONTROL.Cells(Nolig, 1).VALUE = "*" Then
            LigActionControl = Nolig
            For NoCol = 1 To dercol
                If FLCONTROL.Cells(Nolig, NoCol).VALUE = "RESULTAT" Then ColResultatControl = NoCol
            Next
        End If
        If Nolig > LigActionControl And Nolig < 30 Then
            FLCONTROL.Cells(Nolig, COLOK).VALUE = ""
            FLCONTROL.Cells(Nolig, COLRESULTAT).VALUE = ""
            FLCONTROL.Cells(Nolig, COLRESULTATNB).VALUE = ""
            FLCONTROL.Cells(Nolig, COLINPUTNB).VALUE = ""
            FLCONTROL.Cells(Nolig, COLTIME).VALUE = ""
            FLCONTROL.Cells(Nolig, 2).Interior.ColorIndex = 0
        End If
    Next
    
    Call DebutEtape(Etape)
    Call EcritureInput(Etape, "5")
    ' Constitution de la liste des fonctions réservées
    ReDim LISTFCT(1 To 2)
    LISTFCT(1) = "["
    LISTFCT(2) = "interpolation"
    ' Effacement des feuilles intermédiaires et résultats
    NbFile = 1
    Call DelNomenclature
    NbFile = NbFile + 1
    Call DelEquations
    NbFile = NbFile + 1
    FLHMA.Cells.Clear
    NbFile = NbFile + 1
    FLSAI.Cells.Clear
    NbFile = NbFile + 1
    FLCAL.Cells.Clear
    Call EcritureResultats(Etape, "", "" & NbFile)
    Exit Sub
errorHandler: Call ErrorToDo(Etape, "", "" & NbFin, Err)
End Sub
Sub ErrorToDo(Etape As Integer, nom As String, NOMBRE As String, erreur As Variant)
    POURSUIVRE = False
    Call EcritureResultats(Etape, nom, "" & NOMBRE)
    'indique le numéro et la description de l'erreur survenue
    For Nolig = LigActionControl To DerLigControl
        If FLCONTROL.Cells(Nolig, 1) = Etape Then
            FLCONTROL.Cells(Nolig, COLOK) = "ko"
            FLCONTROL.Cells(Nolig, 2).Interior.ColorIndex = 3
        End If
    Next
    MsgBox "Error " & Err.Number & vbLf & Err.Description
End Sub
Function DecAlph(c As Integer) As String
'   =SI(A1<703;SI(A1>26;CAR(ENT((A1-1)/26)+64);"")&SI(A1;CAR(MOD(A1-1;26)+65);"");"")
   DecAlph = IIf(c < 703, IIf(c > 26, Chr((c - 1) \ 26 + 64), "") & _
        IIf(c, Chr(((c - 1) Mod 26) + 65), ""), "")
End Function
Function GetChildren(Classe As String) As String()
'   Retourne les enfants de la classe Energie>EACH retourne Energie>électricité etc
    Dim Retour() As String
    ClasseSplit = Split(Classe, ">EACH")
    Dim avant As String
    avant = ClasseSplit(0)
    Children = GetChildrenClasse(avant)
    ReDim Retour(1 To UBound(Children))
    For No = 1 To UBound(Children)
        Retour(No) = avant & ">" & Children(No)
    Next
    GetChildren = Retour()
End Function
Function GetChildrenClasse(Classe As String) As String()
'   Retourne les enfants de la classe Energie = retourne électricité etc
    Dim Retour() As String
    Dim NoClasse As Integer
    NoClasse = GetPosItemInList(Classe, CLASSES)
    If NoClasse = -1 Then
        MsgBox "La classe " & Classe & " n'existe pas"
        GetChildrenClasse = Retour()
    Else
        GetChildrenClasse = CLASSESI(NoClasse)
    End If
End Function
Function normalisationStringCasse(str As String) As String
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    str = Trim(str)
    str = Replace(str, "é", "e")
    str = Replace(str, "è", "e")
    str = Replace(str, "ô", "o")
    str = Replace(str, "à", "a")
    str = Replace(str, "â", "a")
    str = Replace(str, "ù", "u")
    str = Replace(str, "î", "i")
    str = Replace(str, "ô", "o")
    str = Replace(str, "-", " ")
    str = Replace(str, "(", " ")
    str = Replace(str, ")", " ")
    str = Replace(str, "[", " ")
    str = Replace(str, "]", " ")
    str = Replace(str, "{", " ")
    str = Replace(str, "}", " ")
    str = Replace(str, ",", " ")
    str = Replace(str, ";", " ")
    str = Replace(str, ":", " ")
    str = Replace(str, "!", " ")
    str = Replace(str, "?", " ")
    str = Replace(str, ".", " ")
    str = Replace(str, "/", " ")
    str = Replace(str, "§", " ")
    str = Replace(str, "_", " ")
    str = UCase(Replace(str, "'", " "))
    reg.Pattern = "[ ]+"
    If reg.test(str) Then
        str = reg.Replace(str, " ")
    End If
    reg.Pattern = "#[^#]*#"
    str = str & "#"
    If reg.test(str) Then
        str = reg.Replace(str, "")
    End If
    str = Replace(str, "#", "")
    normalisationStringCasse = UCase(str)
End Function
Function normalisationStringOther(str As String) As String
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    str = Trim(UCase(str))
    str = " " & str & " "
    If Len(str) > 5 Then str = Replace(str, "S ", " ")
    str = Replace(str, " EN ", " ")
    str = Replace(str, " A ", " ")
    str = Replace(str, " AU ", " ")
    str = Replace(str, " AUX ", " ")
    str = Replace(str, " PAR ", " ")
    str = Replace(str, " OU ", " ")
    str = Replace(str, " SOUS ", " ")
    str = Replace(str, " SUR ", " ")
    str = Replace(str, " LA ", " ")
    str = Replace(str, " LE ", " ")
    str = Replace(str, " UN ", " ")
    str = Replace(str, " UNE ", " ")
    str = Replace(str, " LES ", " ")
    str = Replace(str, " DES ", " ")
    str = Replace(str, "ELECTRIQUE ", "ELECTRICITE")
    reg.Pattern = "[ ]+"
    If reg.test(str) Then
        str = reg.Replace(str, " ")
    End If
    str = Trim(str)
    normalisationStringOther = str
End Function
Function resolveFunctions(fct As String) As String
    ' Function sommeproduit
    Dim str As String
    str = fct
    While InStr(str, "sommeproduit") > 0
        ' Détermination de l'argument
        splstrall = Split(str, "sommeproduit")
        splstr = splstrall(1)
        l = 0
        nbpo = 0
        nbpf = 0
        nbco = 1
        nbcf = 0
        For I = 1 To Len(splstr)
            If Mid(splstr, I, 1) = "(" Then
                nbpo = nbpo + 1
            End If
            If Mid(splstr, I, 1) = ")" Then
                nbpf = nbpf + 1
                nbcf = I
            End If
            If nbpo = nbpf Then
                Exit For
            End If
        Next I
        argu = Mid(splstr, 2, nbcf - 2)
        spla = Split(argu, ";")
        newstr = ""
        fin = (UBound(spla) - 1) / 2
        For I = LBound(spla) To fin
            newstr = newstr & "+" & spla(I) & "*" & spla((UBound(spla) + 1) / 2 + I)
        Next I
        newstr = Mid(newstr, 2)
        ' Remplacement du sommeprod
        sp = "sommeproduit(" & argu & ")"
        str = Replace(str, sp, "(" & newstr & ")")
    Wend
    resolveFunctions = str
End Function


Sub setSynonymes()
    openFileIfNot ("MODELE")
    ReDim g_Synonymes(0)
    Dim FLNOM As Worksheet
    Set FLNOM = g_WB_Modele.Worksheets(g_NOMENCLATURE)
    derlignom = getDerLig(FLNOM)
    Dim shortnames() As Variant
    Dim Nshortnames() As Variant
    Dim Nnames() As Variant
    Dim Nsynonymes() As Variant
    Nshortnames = FLNOM.Range("d2:d" & derlignom).VALUE
    Nnames = FLNOM.Range("e2:e" & derlignom).VALUE
    Nsynonymes = FLNOM.Range("g2:g" & derlignom).VALUE
    nl = 0
    For I = LBound(Nshortnames, 1) To UBound(Nshortnames, 1)
        If Trim(Nshortnames(I, 1)) <> "" And Trim(Nsynonymes(I, 1)) <> "" Then
            nl = nl + 1
        End If
    Next
    If nl > 0 Then ReDim g_Synonymes(1 To nl)
    nl = 0
    Dim strSyn As String
    Dim strSho As String
    Dim spl() As String
    Dim splitem As String
    For I = LBound(Nshortnames, 1) To UBound(Nshortnames, 1)
        If Trim(Nshortnames(I, 1)) <> "" And Trim(Nsynonymes(I, 1)) <> "" Then
            nl = nl + 1
            strSyn = Nsynonymes(I, 1)
            strSho = Nshortnames(I, 1)
            spl = Split(strSyn, ",")
            g_Synonymes(nl) = normalisationStringOther(normalisationStringCasse(strSho))
            For j = LBound(spl) To UBound(spl)
                splitem = spl(j)
                g_Synonymes(nl) = g_Synonymes(nl) & "," & normalisationStringOther(normalisationStringCasse(splitem))
            Next
            g_Synonymes(nl) = g_Synonymes(nl) & ","
        End If
    Next
End Sub

Sub ToutesAlaFois()
    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False
    POURSUIVRE = True
    If POURSUIVRE Then Call initialisation
    If POURSUIVRE Then Call SetNomenclature
    If POURSUIVRE Then Call SetEquations
    
    If POURSUIVRE Then Call StructurationNomenclatureHM
    If POURSUIVRE Then Call StructurationTimeHM
    If POURSUIVRE Then Call AlimentationOldDataHM
    If POURSUIVRE Then Call MiseAuFormatHM
    
    If POURSUIVRE Then Call StructurationNomenclatureSaisie
    If POURSUIVRE Then Call StructurationTimeSaisie
    If POURSUIVRE Then Call AlimentationOldDataSaisie
    If POURSUIVRE Then Call MiseAuFormatSaisie
    
    If POURSUIVRE Then Call StructurationNomenclatureCalcul
    If POURSUIVRE Then Call StructurationTimeCalcul
    If POURSUIVRE Then Call AlimentationValeurs
    If POURSUIVRE Then Call AlimentationFormules
    If POURSUIVRE Then Call MiseAuFormatCalcul
    If POURSUIVRE Then Call Finalisation
    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
End Sub
Sub ToutesHM()
    POURSUIVRE = True
    If POURSUIVRE Then Call StructurationNomenclatureHM
    If POURSUIVRE Then Call StructurationTimeHM
    If POURSUIVRE Then Call AlimentationOldDataHM
    If POURSUIVRE Then Call MiseAuFormatHM
End Sub
Sub ToutesTERTIAIRE()
    POURSUIVRE = True
    If POURSUIVRE Then Call StructurationNomenclatureSaisie
    If POURSUIVRE Then Call StructurationTimeSaisie
    If POURSUIVRE Then Call AlimentationOldDataSaisie
    If POURSUIVRE Then Call MiseAuFormatSaisie
End Sub
Sub ToutesCALCULS()
    POURSUIVRE = True
    If POURSUIVRE Then Call StructurationNomenclatureCalcul
    If POURSUIVRE Then Call StructurationTimeCalcul
    If POURSUIVRE Then Call AlimentationValeurs
    If POURSUIVRE Then Call AlimentationFormules
    If POURSUIVRE Then Call MiseAuFormatCalcul
End Sub
Sub SetFl()
    Set FLCONTROL = Worksheets("CONTROL")
    Set FL0NO = Worksheets("0NOMENCLATURE")
    Set FLNOM = Worksheets("NOMENCLATURE")
    Set FL0EQ = Worksheets("0EQUATIONS")
    Set FLEQU = Worksheets("EQUATIONS")
    Set FLTIM = Worksheets("0TIME")
    Set FLSSK = Worksheets("0TERTIAIRE")
    Set FLSAI = Worksheets("TERTIAIRE")
    Set FLCSK = Worksheets("0CALCULS")
    Set FLCAL = Worksheets("calculs TERTIAIRE")
    Set FLOLD = Worksheets("PRECEDENTTERTIAIRE")
    Set FLHMA = Worksheets("Hypothèses Macro")
    Set FLOHM = Worksheets("PRECEDENTHM")
    Set FL0HM = Worksheets("0HM")
End Sub
Sub AlimentationFormules()
    Etape = 15
    Call DebutEtape(Etape)
    ''''On Error GoTo ErrorHandler
    a0 = InitTime()
    Call SetFl
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Dim PosFct As Integer
    Dim LigFct As Integer
    NOMBRE = 0
    NbEquations = 0
    derlig = Split(FLCAL.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    LigCal = GetFirstLine(FLCAL)
    ColCal = GetFirstCol(FLCAL)
    PosFct = ColCal(1)
    FORMULESLOCALES = FLCAL.Range("A1:" & DecAlph(PosFct) & LigCal(1)).FormulaLocal
    Dim fct As String
    a1 = SetTime()
    'ReDim QUANTITECAL(1 To LigCal(1))
    'MsgBox DecAlph(PosFct) & "1:" & DecAlph(PosFct) & LigCal(1)
    QUANTITECAL = FLCAL.Range(DecAlph(PosFct) & "1:" & DecAlph(PosFct) & LigCal(1)).VALUE
    'MsgBox LBound(QUANTITECAL) & ":" & UBound(QUANTITECAL) & (LigCal(1) - 1)
    a2 = SetTime()
    For Nolig = 1 To LigCal(1) - 1
        '''Var1 = FLCAL.Cells(NoLig, PosFct)
        Var1 = QUANTITECAL(Nolig, 1)
        If Left(Var1, 1) = "[" Then
            fct = Split(Var1, "°")(0)
            LigFct = Nolig
            NbEquations = NbEquations + 1
            Call ChercherFormules(fct, LigFct)
            'a1 = SetTime()
        End If
    Next
    a3 = SetTime()
    FLCAL.Range("A1:" & DecAlph(PosFct) & LigCal(1)).FormulaLocal = FORMULESLOCALES
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call EcritureInput(Etape, "" & NbEquations)
    'retour à la feuille de contrôle
    
    Call EcritureResultats(Etape, "calculs TERTIAIRE", "" & NOMBRE)
    'MsgBox GetTime()
    Exit Sub
errorHandler: Call ErrorToDo(Etape, "calculs TERTIAIRE", "" & NOMBRE, Err)
End Sub
Function GetPosFctInCal(FctToFind As String, FLName As String) As Integer
' Retourne la position ligne d'une fonction Fct dans la feuille de calculs ou Macro
    FctToFind = Split(FctToFind, "(")(0)
    Set FL = Worksheets(FLName)
    derlig = Split(FL.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To dercol
        var = FL.Cells(1, NoCol)
        If Left(var, 3) = "END" Then PosFct = NoCol
    Next
    Dim LigFct As Integer
    LigFct = 0
    For Nolig = 1 To derlig - 1
        Var1 = FL.Cells(Nolig, PosFct).VALUE
        If Left(Var1, 1) = "[" Then
            fct = Split(Split(Var1, "°")(0), "(")(0)
            If FctToFind = fct Then
                LigFct = Nolig
                Exit For
            End If
        End If
    Next
    GetPosFctInCal = LigFct
End Function
Function GetFeuFctIn(FctToFind As String) As String
' Retourne le nom de la feuille d'une fonction Fct
    Dim Ret As String
    Ret = ""
    FctToFind = Split(FctToFind, "(")(0)
    derlig = Split(FLCAL.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To dercol
        var = FLCAL.Cells(1, NoCol)
        If Left(var, 3) = "END" Then PosFct = NoCol
    Next
    Dim LigFct As Integer
    LigFct = 0
    For Nolig = 1 To derlig - 1
        Var1 = FLCAL.Cells(Nolig, PosFct).VALUE
        If Left(Var1, 1) = "[" Then
            fct = Split(Split(Var1, "°")(0), "(")(0)
            If FctToFind = fct Then
                LigFct = Nolig
                Exit For
            End If
        End If
    Next

    If LigFct > 0 Then Ret = FLCAL.NAME
    derlig = Split(FLHMA.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLHMA.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To dercol
        var = FLHMA.Cells(1, NoCol)
        If Left(var, 3) = "END" Then PosFct = NoCol
    Next
    LigFct = 0
    For Nolig = 1 To derlig - 1
        Var1 = FLHMA.Cells(Nolig, PosFct).VALUE
        If Left(Var1, 1) = "[" Then
            fct = Split(Split(Var1, "°")(0), "(")(0)
            If FctToFind = fct Then
                LigFct = Nolig
                Exit For
            End If
        End If
    Next
    If LigFct > 0 Then Ret = FLHMA.NAME
    GetFeuFctIn = Ret
End Function
Function GetFctJalInEqu(fct As String) As String
' Get de la feuille et formule de la quantité Fct pour les dates jalons
    derlig = Split(FLEQU.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLEQU.UsedRange.Address, "$")(3)).Column
    Dim TextFct As String
    TextFct = ""
    Dim LigFct As Integer
    LigFct = 0
    For Nolig = 1 To derlig
        var = FLEQU.Cells(Nolig, 1)
        If Left(var, 1) = "*" Then
            DebFet = Nolig + 1
            Exit For
        End If
    Next
    Dim res() As String
    For Nolig = DebFet To derlig
        FctLig = "[" & FLEQU.Cells(Nolig, 4) & "]." & FLEQU.Cells(Nolig, 5) & "(t)"
        If fct = FctLig Then
            LigFct = Nolig
            If FLEQU.Cells(Nolig, 11).VALUE <> "" Then TextFct = FLEQU.Cells(Nolig, 11).VALUE & "'" & FLEQU.Cells(Nolig, 12).VALUE
        End If
    Next
    GetFctJalInEqu = TextFct
End Function
Function GetFctHisInEqu(fct As String) As String
' Get de la feuille et formule de la quantité Fct pour l'historique
    derlig = Split(FLEQU.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLEQU.UsedRange.Address, "$")(3)).Column
    Dim TextFct As String
    TextFct = ""
    For Nolig = 1 To derlig
        var = FLEQU.Cells(Nolig, 1)
        If Left(var, 1) = "*" Then
            DebFet = Nolig + 1
            Exit For
        End If
    Next
    Dim res() As String
    For Nolig = DebFet To derlig
        FctLig = "[" & FLEQU.Cells(Nolig, 4) & "]." & FLEQU.Cells(Nolig, 5) & "(t)"
        If fct = FctLig Then
            ' Analyse et transformation de l'équation
            If FLEQU.Cells(Nolig, 9).VALUE <> "" Then TextFct = FLEQU.Cells(Nolig, 9).VALUE & "'" & FLEQU.Cells(Nolig, 10).VALUE
        End If
    Next
    GetFctHisInEqu = TextFct
End Function
Sub ChercherFormules(fct As String, PosFct As Integer)
' Get de l'équation associée à la quantité
    Dim TextFct As String
    For Nolig = 1 To UBound(EQUATIONS)
        FctLig = "[" & EQUATIONS(Nolig) & "]." & EQUATIONSQ(Nolig) & "(t)"
        If fct = FctLig Then
            TextFct = EQUATIONSF(Nolig)
            Call ProcessFct(TextFct, PosFct)
        End If
    Next
End Sub
Sub ProcessFct(TextFctToProcess As String, PosFct As Integer)
' Retourne la liste des fonctions de la chaine de caractères
    NoFctIn = 0
    Dim Retour As String
    Dim Item() As String
    Dim Formules() As Variant
    Dim TestText As String
    Dim CurrentTime As Integer
    CurrentTime = 0
    TestText = Left(TextFctToProcess, 13)
    For NoF = 1 To 50
        PosDeb = 0
        For NoFct = LBound(LISTFCT) To UBound(LISTFCT)
            PosDeb = InStr(TextFctToProcess, LISTFCT(NoFct))
            If PosDeb > 0 Then
                NoFctIn = NoFctIn + 1
                PosFin = InStr(PosDeb, TextFctToProcess, ")")
                Retour = Mid(TextFctToProcess, PosDeb, PosFin - PosDeb + 1)
                ReDim Preserve Item(1 To NoFctIn)
                Item(NoFctIn) = Retour
                TextFctToProcess = Replace(TextFctToProcess, Retour, "FUNCTION" & NoFctIn)
                Exit For
            End If
        Next
        If PosDeb > 0 Then
            ReDim Preserve Formules(1 To NoFctIn)
            Dim PosFctFinal As Integer
            Dim SourceFL As String
            If Left(Retour, 1) = "[" Then
                SourceFL = GetFeuFctIn(Item(NoFctIn))
                PosFctFinal = GetPosFctInCal(Item(NoFctIn), SourceFL)
                If Right(Split(Retour, ")")(0), 3) = "t-1" Then CurrentTime = -1
                If Right(Split(Retour, ")")(0), 1) = "t" Then CurrentTime = 0
            End If
            If Left(Retour, 13) = "interpolation" Then
                SourceFL = FLCAL.NAME
                PosFctFinal = PosFct
                CurrentTime = 0
            End If
            ' Calcul de la position de la quantité cherchée
            'If TestText <> "interpolation" Then
                'aaa = GetDataFct(PosFctFinal)
            'End If
            Formules(NoFctIn) = AnalyseEtRemplacement(SourceFL, PosFctFinal, Retour, CurrentTime)
        Else
            Exit For
        End If
    Next

    FctData = GetDataFct(PosFct)
    Dim ResDataFct() As String
    For NoData = LBound(FctData) To UBound(FctData)
        ReDim Preserve ResDataFct(LBound(FctData) To NoData)
        TFTP = TextFctToProcess
        SplFctData = Split(FctData(NoData), "°")
        If NoFctIn > 0 Then
            ' Replace des FUNCTIONn par les formules
            For NoFctToReplace = 1 To NoFctIn
                ' entre parenthèses???
                spl = Split(Formules(NoFctToReplace)(NoData), "°")
                Target = spl(1)
                TFTP = Replace(TFTP, "FUNCTION" & NoFctToReplace, Target)
            Next
            ResDataFct(NoData) = SplFctData(0) & "°" & TFTP
        Else
            ResDataFct(NoData) = FctData(NoData)
        End If
    Next
    Call MajFormules(ResDataFct, PosFct)
    ' Puis alimentation des cellules
End Sub
Function AnalyseEtRemplacement(FLName As String, NoLigFct As Integer, equa As String, CurT As Integer) As String()
    Set FL = Worksheets(FLName)
    Dim Vect() As String
    Vect = GetDataFct(NoLigFct)
    Dim res() As String
    ReDim res(1 To 1)
    res(1) = "KO"
    If equa = "interpolation(t)" Then
        res = interpolation(Vect, NoLigFct)
    End If
'If UBound(Vect) < 1 Then MsgBox NoLigFct & ":" & Equa & "::" & UBound(Vect)
    If Left(equa, 1) = "[" Then
        For NoVal = LBound(Vect) To UBound(Vect)
            spl = Split(Vect(NoVal), "°")
            ReDim Preserve res(1 To NoVal)
            res(NoVal) = spl(0) & "°" & "'" & FL.NAME & "'!" & DecAlph(Val(spl(2)) + CurT) & NoLigFct & "°" & spl(2)
            'Split(ListVa(NoVal), "°")(0) & "°" & calc & "°" & Split(ListVa(NoVal), "°")(2)
            'ValVal = Split(Vect(NoVal), "°")(1)
        Next
    
        'Res = Vect 'GetDataFct(GetLigFctInCal(Equa))
    End If
    
    AnalyseEtRemplacement = res
End Function
Sub MajFormules(Valeurs() As String, NoLigVal As Integer)
' Mise à jour des formules dans l'onglet calculs
    DerLigCal = Split(FLCAL.UsedRange.Address, "$")(4)
    DerColCal = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    For Nolig = 1 To DerLigCal
        var = FLCAL.Cells(Nolig, 1)
        If Left(var, 3) = "END" Then ENDLIG = Nolig
    Next
    ''A0 = InitTime()
    For NoCol = 1 To DerColCal
        IdLig = Split(FLCAL.Cells(ENDLIG, NoCol), "°")(0)
        For NoVal = LBound(Valeurs) To UBound(Valeurs)
            IdVal = Split(Valeurs(NoVal), "°")(0)
            If IdLig = IdVal Then
                Dim ValVal As String
                ValVal = Split(Valeurs(NoVal), "°")(1)
                If ValVal <> "" Then
                    TypeTime = Split(IdLig, "$")(0)
                    'If TypeTime = "h" Or TypeTime = "j" Then
                        'FLCAL.Cells(NoLigVal, NoCol).Value = Val(ValVal)
                    'End If FORMULESLOCALES
                    ''''If FLCAL.Cells(NoLigVal, NoCol).Formula = "" Or Left(FLCAL.Cells(NoLigVal, NoCol), 1) = "[" Then
                        ''''ValVal = Replace(ValVal, "'" & FLCAL.Name & "'!", "")
                        ''''FLCAL.Cells(NoLigVal, NoCol).FormulaLocal = "=" & ValVal
                        ''''NOMBRE = NOMBRE + 1
                    ''''End If
                    If FORMULESLOCALES(NoLigVal, NoCol) = "" Or Left(FLCAL.Cells(NoLigVal, NoCol), 1) = "[" Then
                        ValVal = Replace(ValVal, "'" & FLCAL.NAME & "'!", "")
                        FORMULESLOCALES(NoLigVal, NoCol) = "=" & ValVal
                        NOMBRE = NOMBRE + 1
                    End If
                End If
                'ValVal = Replace(ValVal, ",", ".")
                'If ValVal <> "" Then
                   ' FLCAL.Cells(NoLigVal, NoCol).Value = Val(ValVal)
                    'FLCAL.Cells(NoLigVal, NoCol).Formula = ValVal
                'End If
                'Double.Parse(ValVal)
                'FLCAL.Cells(NoLigVal, NoCol) = CDbl(Val(ValVal))
                'ValVal = Replace(ValVal, ",", ".")
                'FLCAL.Cells(NoLigVal, NoCol) = ToDouble(ValVal)
            End If
        Next
    Next
    ''MsgBox "MajFormules " & GetTime()
End Sub
Function GetLigFctInCal(fct As String) As Integer
    DerLigCal = Split(FLCAL.UsedRange.Address, "$")(4)
    DerColCal = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To DerColCal
        var = FLCAL.Cells(1, NoCol)
        If Left(var, 3) = "END" Then ENDCOL = NoCol
    Next
    For Nolig = 1 To DerLigCal
        var = FLCAL.Cells(Nolig, 1)
        If Left(var, 3) = "END" Then ENDLIG = Nolig
    Next
    Dim Retour As Integer
    Retour = 0
    For Nolig = 1 To ENDLIG - 1
        var = FLCAL.Cells(Nolig, ENDCOL)
        spl = Split(var, "°")
        If UBound(spl) > 0 Then
            If spl(0) = fct Then
                Retour = Nolig
                Exit For
            End If
        End If
    Next
    GetLigFctInCal = Retour
End Function
Function GetDataFct(NoLigFct As Integer) As String()
    DerLigCal = Split(FLCAL.UsedRange.Address, "$")(4)
    DerColCal = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To DerColCal
        var = FLCAL.Cells(1, NoCol)
        If Left(var, 3) = "END" Then ENDCOL = NoCol
    Next
    For Nolig = 1 To DerLigCal
        var = FLCAL.Cells(Nolig, 1)
        If Left(var, 3) = "END" Then ENDLIG = Nolig
    Next
    Dim EndColCal As Integer
    EndColCal = ENDCOL - 1
    Dim dates() As String
    ReDim dates(1 To EndColCal)
    Dim Retour() As String
    ReDim Retour(1 To 1)
    Dim Cmpt As Integer
    Cmpt = 1
    Dim spl() As String
    Dim SplDates() As String
    On Error GoTo errorHandler
    For NoCol = 1 To ENDCOL - 1
        dates(NoCol) = FLCAL.Cells(ENDLIG, NoCol)
        spl() = Split(dates(NoCol), "$")
        SplDates() = Split(dates(NoCol), "°")
'If NoCol >= 28 Then MsgBox NoCol & ":" & Dates(NoCol - 1) & ":" & Dates(NoCol)
        If UBound(spl) > 0 Then
            ReDim Preserve Retour(1 To Cmpt)
            If spl(0) = "h" Or spl(0) = "j" Then
'If NoCol >= 28 Then
    'MsgBox NoCol & ":" & SplDates(0) & ":" & Spl(0)
    'MsgBox NoCol & ":" & Cmpt & ":" & FLCAL.Cells(NoLigFct, NoCol).Value
'End If
'If NoCol >= 28 Then MsgBox NoCol & ":" & Cmpt & ""
                Retour(Cmpt) = SplDates(0) & "°" & FLCAL.Cells(NoLigFct, NoCol).VALUE & "°" & NoCol
'If NoCol >= 28 Then MsgBox NoCol & ":" & Cmpt & ":" & Retour(Cmpt)
            Else
                Retour(Cmpt) = SplDates(0) & "°" & "°" & NoCol
            End If
            Cmpt = Cmpt + 1
        End If
    Next
    GetDataFct = Retour()
    Exit Function
errorHandler:
    MsgBox "err " & NoLigFct & ":" & Join(dates, " ")
    MsgBox "err " & NoCol & "=" & (ENDCOL - 1) & "<" & UBound(dates) & "=" & EndColCal
    'MsgBox FLCAL.Cells(NoLigFct, NoCol).Value
    
End Function
Function interpolation(ListVa() As String, NoLigFct As Integer) As String()
    ' En valeurs pour commencer
    Dim Retour() As String
    Dim ValCal As Double
    ReDim Retour(LBound(ListVa) To UBound(ListVa))
    Dim LastValKnown As String
    LastValKnown = ""
    Dim NextValKnown As String
    NextValKnown = ""
    Dim NoValLast As Integer
    Dim NoValNext As Integer
    Dim NoColLast As Integer
    Dim NoColNext As Integer
    NoValLast = 0
    NoValNext = 0
    Dim NoCol As Integer
    For NoVal = LBound(ListVa) To UBound(ListVa)
        ValVal = Split(ListVa(NoVal), "°")(1)
        If ValVal = "" Then
            NoCol = Val(Split(ListVa(NoVal), "°")(2))
            For NoValN = NoVal + 1 To UBound(ListVa)
                SuiteVal = Split(ListVa(NoValN), "°")(1)
                If SuiteVal <> "" Then
                    NextValKnown = Split(ListVa(NoValN), "°")(1)
                    NextValKnown = Replace(NextValKnown, ",", ".")
                    NoValNext = NoValN
                    NoColNext = Val(Split(ListVa(NoValN), "°")(2))
                    Exit For
                End If
            Next
            ' voir Novaldeb??? =TERTIAIRE!C30

            calc = "('calculs TERTIAIRE'!" & DecAlph(NoCol - 1) & NoLigFct
            calc = calc & "+ (" & "'calculs TERTIAIRE'!" & DecAlph(NoColNext) & NoLigFct
            calc = calc & " - " & "'calculs TERTIAIRE'!" & DecAlph(NoColLast) & NoLigFct & ")"
            calc = calc & "/ (" & NoValNext & " - " & NoValLast & "))"
            'Val(Split(ListVa(NoVal), "°")(1)) + (Val(NextValKnown) - Val(LastValKnown)) / (NoValNext - NoValLast)
            Retour(NoVal) = Split(ListVa(NoVal), "°")(0) & "°" & calc & "°" & Split(ListVa(NoVal), "°")(2)
        Else
            LastValKnown = Split(ListVa(NoVal), "°")(1)
            LastValKnown = Replace(LastValKnown, ",", ".")
            NoValLast = NoVal
            NoColLast = Val(Split(ListVa(NoVal), "°")(2))
            Retour(NoVal) = ListVa(NoVal)
        End If
    Next
    interpolation = Retour()
End Function
Sub Finalisation()
    Etape = 17
    Call DebutEtape(Etape)
    Call SetFl
    Call FinalisationFL(FLHMA)
    Call FinalisationFL(FLSAI)
    Call FinalisationFL(FLCAL)
    Call EcritureInput(Etape, "3")
    Call EcritureResultats(Etape, "", "3")
End Sub
Sub FinalisationFL(FLOUT As Worksheet)
    derlig = Split(FLOUT.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLOUT.UsedRange.Address, "$")(3)).Column
    LigEndSAI = 0
    For Nolig = 1 To derlig
        var = FLOUT.Cells(Nolig, 1).VALUE
        If Left(var, 3) = "END" Then
            LigEndSAI = Nolig
            Exit For
        End If
    Next
    ColEndSAI = 0
    For NoCol = 1 To dercol
        var = FLOUT.Cells(1, NoCol).VALUE
        If Left(var, 3) = "END" Then
            ColEndSAI = NoCol
            Exit For
        End If
    Next
    If LigEndSAI > 0 Then
        For Nolig = LigEndSAI To derlig
            For NoCol = 1 To dercol
                FLOUT.Cells(Nolig, NoCol).VALUE = ""
            Next
        Next
    End If
    If ColEndSAI > 0 Then
        For NoCol = ColEndSAI To dercol
            For Nolig = 1 To derlig
                FLOUT.Cells(Nolig, NoCol).VALUE = ""
            Next
        Next
    End If
End Sub
Sub DelNomenclature()
    derlig = Split(FLNOM.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLNOM.Cells(Nolig, 1) = "*" Then
            LigDebNom = Nolig + 1
            Exit For
        End If
    Next
    'For NoLig = DerLig To LigDebNom Step -1
        'FLNOM.Rows(NoLig).Delete
    'Next
    FLNOM.Rows(LigDebNom & ":" & FLNOM.Rows.Count).Clear
End Sub
Sub DelHM()
    spl = Split(FLHMA.UsedRange.Address, "$")
    If UBound(spl) > 3 Then
        derlig = Split(FLHMA.UsedRange.Address, "$")(4)
        For Nolig = derlig To 1 Step -1
            FLHMA.Rows(Nolig).Delete Shift:=xlUp
        Next
    End If
End Sub
Sub DelEquations()
    derlig = Split(FLEQU.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLEQU.Cells(Nolig, 1) = "*" Then
            LigDebEqu = Nolig + 1
            Exit For
        End If
    Next
    FLEQU.Rows(LigDebEqu & ":" & FLEQU.Rows.Count).Clear
End Sub
Function GetPosItemInList(ToFind As String, list() As String) As Double
'   renvoie la position de ToFind dans List et -1 sinon
    On Error GoTo errorHandler
    Dim NoItems As Double
    NoItems = Application.WorksheetFunction.Match(ToFind, list, 0)
    GetPosItemInList = NoItems
    Exit Function
errorHandler:
    MsgBox ("erreur GetPosItemInList>" & GetPosItemInList & "<>" & ToFind & "<in>" & Join(list, ";") & "<")
    GetPosItemInList = -1
End Function
Function nettoyage(anettoyer As Variant) As String
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.MultiLine = False
    reg.IgnoreCase = False
    reg.Global = True
    'reg.Pattern = ":$"
    'nettoyage = reg.Replace(anettoyer, "")
    'reg.Pattern = ">$"
    'nettoyage = reg.Replace(nettoyage, "")
    reg.Pattern = "[ ]*,[ ]*"
    nettoyage = reg.Replace(anettoyer, ",")
    'reg.Pattern = "(.>)?\.$"
    'nettoyage = reg.Replace(nettoyage, "")
End Function
Function remplaceDoubleSupArray(ToReplace() As String, entree() As String) As String()
    Dim Nolig As Integer, text As String
    For Nolig = LBound(ToReplace, 1) To UBound(ToReplace, 1)
        text = ToReplace(Nolig, 1)
        text = remplaceDoubleSup(text, entree)
    Next
    remplaceDoubleSupArray = ToReplace
End Function
Function remplaceDoubleSupArraySimples(entree() As String) As String()
    Dim NoOne As Integer
    For NoOne = LBound(entree, 1) To UBound(entree, 1)
        entree(NoOne, 1) = remplaceDoubleSup(entree(NoOne, 1), entree)
        If entree(NoOne, 1) Like "*>>*" Then
Dim newLogLine() As Variant
'MsgBox "ici"
newLogLine = Array("", "ERREUR", "0NOM", "", "", time(), "", entree(NoOne, 1), "", "raccourci non résolu")
logNom = alimLog(logNom, newLogLine)
        End If
    Next
    remplaceDoubleSupArraySimples = entree
End Function
Function remplaceDoubleSup(remplace As String, entree() As String, ORIGINE As String, numlig As Integer) As String
' passer d'autres arguments pour alerte raccourci multiple
    Dim ok As String, newremplace As String
    Dim avant As String, apres As String
    ok = ""
    Dim listOk() As String
    splr = Split(remplace, ".")
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    Dim rempfin As String
    Dim jct As String
    jct = ""
    rempfin = ""
    nbr = 0
    Dim where As String
    where = ""
    Dim Last As String
    Dim numf As Integer
    numf = 0
    Dim avantS As String
    Dim apresS As String

'
'If numlig = 41 Then MsgBox remplace
'MsgBox remplace & Chr(10) & Join(splr, Chr(10))
    For Each remp In splr
        avantS = ""
        apresS = ""
        If InStr(remp, "[") > 0 Then
            avantS = Left(remp, InStr(remp, "["))
            remp = Mid(remp, InStr(remp, "[") + 1)
        End If
        If InStr(remp, "]") > 0 Then
            apresS = Mid(remp, InStr(remp, "]"))
            remp = Mid(remp, 1, InStr(remp, "]") - 1)
        End If
        rempres = remp
        numf = numf + 1
        splRemp = Split(remp, ":")
        If UBound(splRemp) > 0 Then
            att = ":" & splRemp(1)
        Else
            att = ""
        End If
        If remp Like "*>>*" Then
            remp = splRemp(0)
'If numlig = 127 Then MsgBox remp & Chr(10) & att
            spl = Split(remp, ">")
            avant = ""
            apres = ""
            Last = ">"
            ReDim listOk(0)
            For Nolig = LBound(spl) To UBound(spl)
                If spl(Nolig) = "" And Last <> ">" Then
                    apres = spl(Nolig + 1)
                    avant = spl(Nolig - 1)
                    For noent = LBound(entree) To UBound(entree)
                        If entree(noent) Like "*.*" Then GoTo continueLoop
                        If remp = entree(noent) Then
                            Exit For
                        End If
                        If Right(avant, 1) = "." Then
                            avant = ""
                        End If
                        reg.Pattern = avant & "(.*)" & apres
                        reg.Global = False
'MsgBox noent & ">>>" & entree(noent) & "<<<" & reg.Pattern
'If numlig > 316 Then MsgBox noent & ">>>" & entree(noent) & "<<<" & reg.Pattern
                        If reg.test(entree(noent)) Then
                            Set matches = reg.Execute(entree(noent))
                            For Each Match In matches
                                ok = Match.SubMatches(0)
'If numLig = 40 Then MsgBox numLig & ":" & ok & ":" & InStr(ok, apres) & ":apres=" & apres & ":right>>>" & Right(ok, 1) & "<<<"
                                tni = 1
                                If InStr(Split(remp, ">>")(1), ">") = 0 Then
                                    ' A VERIFIER
                                    testNotInclude1 = InStr(">" & entree(noent) & ":", ">" & Split(remp, ">>")(1) & ":")
                                    testNotInclude2 = InStr(">" & entree(noent) & ">", ">" & Split(remp, ">>")(1) & ">")
                                    tni = testNotInclude1 + testNotInclude2
'MsgBox entree(noent) & Chr(10) & Split(remp, ">>")(1)
                                End If
'If numlig = 41 Then MsgBox remplace & Chr(10) & ok & ">>>>>" & apres & Chr(10) & entree(noent) & "|" & Split(remp, ">>")(1) & Chr(10) & testNotInclude1 & ":" & testNotInclude2
                                If InStr(ok, apres) = 0 And (Right(ok, 1) = "" Or Right(ok, 1) = ">") And tni > 0 Then
                                'If InStr(ok, apres) = 0 And (Right(ok, 1) = "" Or Right(ok, 1) = ">") Then
                                    rempres = Replace(remp, ">>", ok)
'If numlig = 41 Then MsgBox ">>>>>" & rempres
'If numlig = 127 Then MsgBox remp & Chr(10) & reg.Pattern & Chr(10) & noent & ":" & remp & Chr(10) & "0>" & ok & Chr(10) & entree(noent)
                                    If UBound(listOk) = 0 Then
                                        ReDim listOk(1 To 1)
                                        listOk(1) = ok
                                        'where = Split(entree(noent), ":")(0)
                                        nbr = 1
'If numLig = 40 Then MsgBox numLig & ":" & ok
                                    Else
                                        trouve = False
                                        For I = LBound(listOk) To UBound(listOk)
                                            If ok = listOk(I) Then
                                                trouve = True
                                                Exit For
                                            End If
                                        Next
                                        If Not trouve Then
'If numLig = 40 Then MsgBox numLig & ":" & Join(listOk, "|")
                                            ReDim Preserve listOk(1 To (UBound(listOk) + 1))
                                            listOk(UBound(listOk)) = ok
                                            'where = where & "," & Split(entree(noent), ":")(0)
                                            If where = "" Then where = "BEFORE" & numf
                                            If where <> "" Then
                                                lastWhere = Split(where, ",")(0)
                                                If lastWhere <> ("BEFORE" & numf) Then where = where & "," & "BEFORE" & numf
                                            End If
                                            nbr = nbr + 1
                                        End If
                                    End If
                                End If
                            Next
                        End If
continueLoop:
                    Next
                    Exit For
                End If
                Last = spl(Nolig)
            Next
'If numlig = 127 Then MsgBox remp & Chr(10) & "1>" & rempres
        End If
        'rempres = avantS & rempres & apresS
'MsgBox (UBound(Split(rempres, ":")) > 0) & Chr(10) & remplace & Chr(10) & rempfin & jct & rempres & Chr(10) & rempfin & jct & rempres & att
        If UBound(Split(rempres, ":")) > 0 Then
            rempfin = rempfin & jct & avantS & rempres & apresS
        Else
            rempfin = rempfin & jct & avantS & rempres & att & apresS
        End If
        'If apresS <> "" Then MsgBox rempres & Chr(10) & apresS
'MsgBox ">>>" & rempfin
        'If UBound(Split(rempres, ":")) > 0 Then
'If numlig = 275 Then MsgBox ">>>" & rempres
' pourquoi j'ai enlevé le & att ??????
        '''rempfin = rempfin & jct & rempres
        'Else
            'rempfin = rempfin & jct & rempres & att
        'End If
        jct = "."
        'If nbr > 1 Then
            'where = where & "," & "BEFORE" & numf
            'MsgBox where
        'End If
    Next
'MsgBox remplace & Chr(10) & rempfin
'If numlig = 164 Then MsgBox remplace & Chr(10) & rempfin
'If numlig = 41 Then MsgBox remplace & ":" & nbr & ":" & rempfin
    If nbr > 1 And numlig > 0 Then
        Dim newLogLine() As Variant
newLogLine = Array("", "ALERTE", ORIGINE, "Raccourci", time(), where, remplace, numlig, nbr & " raccourcis possibles")
logNom = alimLog(logNom, newLogLine)
    End If
    remplaceDoubleSup = rempfin
End Function


Function alimLog(entree() As Variant, alim() As Variant) As Variant()
    Dim test As Boolean
    test = False
    Dim als As String
    Dim ens As String
    als = ""
    ens = ""
    For I = LBound(alim) To UBound(alim)
        If I <> 0 And I <> 4 Then
            If I = 1 And Mid(alim(I), 1, 6) = "ERREUR" Then
                als = als & Mid(alim(I), 1, 6)
                ens = ens & Mid(entree(I + 1, UBound(entree, 2)), 1, 6)
            Else
                als = als & alim(I)
                ens = ens & entree(I + 1, UBound(entree, 2))
            End If
        End If
    Next
    Dim nbe As Integer
    If ens <> als Then
        ReDim Preserve entree(1 To UBound(entree, 1), LBound(entree, 2) To UBound(entree, 2) + 1)
        For I = LBound(alim) To UBound(alim)
            entree(I + 1, UBound(entree, 2)) = alim(I)
        Next
    Else
        If Mid(alim(1), 1, 6) = "ERREUR" Then
            entree(2, UBound(entree, 2)) = "ERREUR"
        End If
    End If
    alimLog = entree
End Function
Sub clearLogNom(FL As Worksheet)
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    If FL.Cells(1, 1).VALUE = "" Then
        For ii = LBound(g_listHeadCR) To UBound(g_listHeadCR)
            FL.Cells(1, ii).VALUE = g_listHeadCR(ii)
        Next
    End If
    derligsheet = getDerLig(FL)
    DerColSheet = getDerCol(FL)
    If derligsheet > 1 Then FL.Range("A2:K" & derligsheet).Clear
    With FL.Range("a1:" & DecAlph(DerColSheet) & "1")
        ReDim logNom(1 To DerColSheet, 1 To 1)
        logNom = Application.Transpose(.VALUE)
    End With
End Sub

Function ajoutNbQuaToError(entree() As Variant, listnbq() As String, deb As Integer) As Variant()
    Dim pos As Integer
    For j = LBound(entree, 2) To UBound(entree, 2)
        If Mid(entree(2, j), 1, 6) = "ERREUR" And Len(entree(2, j)) > 6 Then
            pos = CInt(entree(8, j)) - deb + 1
'MsgBox entree(2, j) & Chr(10) & pos & Chr(10) & LBound(listnbq) & ":" & UBound(listnbq)
            If CInt(listnbq(pos, 1)) <> CInt(Mid(entree(2, j), 7, Len(entree(2, j)))) Then
                entree(2, j) = entree(2, j) & "/" & listnbq(pos, 1)
            Else
                entree(2, j) = "ERREUR"
            End If
        End If
    Next
    ajoutNbQuaToError = entree
End Function



Function TestNotAlpha(t As String) As Boolean
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = "[éèàùçâôîïêû0-9a-zA-Z_\-'/ ]+"
    Ret = reg.Replace(t, "")
    If Ret <> "" Then TestNotAlpha = True
    If Ret = "" Then TestNotAlpha = False
End Function
Function TestNotAlphaQua(t As String) As Boolean
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = "[éèàùçâôîïêû0-9a-zA-Z_\-'/ %]+"
    Ret = reg.Replace(t, "")
    If Ret <> "" Then TestNotAlphaQua = True
    If Ret = "" Then TestNotAlphaQua = False
End Function
Function TestNotAlphaVir(t As String) As Boolean
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = "[éèàùçâôîïêû0-9a-zA-Z_\-'/, ]+"
    Ret = reg.Replace(t, "")
    If Ret <> "" Then TestNotAlphaVir = True
    If Ret = "" Then TestNotAlphaVir = False
End Function
Function TestNotAlphaVirOp(t As String) As Boolean
    t = extraitAttribut(t)
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = "[éèàùçâôîïêû0-9a-zA-Z_\-'/, ]+"
    Ret = reg.Replace(t, "")
    If Ret <> "" Then TestNotAlphaVirOp = True
    If Ret = "" Then TestNotAlphaVirOp = False
End Function
Function TestNotAlphaSup(t As String) As Boolean
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = "[éèàùçâôîïêû0-9a-zA-Z_\-'/> ]+"
    Ret = reg.Replace(t, "")
    If Ret <> "" Then TestNotAlphaSup = True
    If Ret = "" Then TestNotAlphaSup = False
End Function
Sub testFilter()
    Dim FLQUA As Worksheet
    Set FLQUA = Worksheets("QUANTITY")
    ' Lecture des quantités
    Dim qua() As Variant
    Dim Quantities() As String
    Dim derlig As Integer
    Dim dercol As Integer
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLQUA.Range("a" & 2 & ":" & DecAlph(dercol) & derlig)
        ReDim qua(1 To (derlig - 1), 1 To dercol)
        qua = .VALUE
    End With
    Dim quaPos() As String
    ReDim quaPos(1 To UBound(qua, 1), 1 To UBound(qua, 2))
    For I = LBound(qua, 1) To UBound(qua, 1)
        For j = LBound(qua, 2) To UBound(qua, 2)
            quaPos(I, j) = qua(I, j)
            '''quaPos(i, j + 1) = ""
        Next
    Next
    ReDim Quantities(1 To UBound(qua, 1))
    For I = LBound(qua, 1) To UBound(qua, 1)
        Quantities(I) = I & "@[" & qua(I, 3) & "][" & qua(I, 4) & "][" & qua(I, 5) & "]." & qua(I, 6)
    Next
    ' fin lecture
    achercher = "USAGES>ECS.TECHNOLOGIES>Gaz:Gaz,Chauffage,ECS.SECTEUR>TERTIAIRE>BRANCHE>Cafés/Hôtels/Restaurants.AGE>RECENT>Neuf][a][s>REPRISE].PdM"
    sngChrono = Timer
    Dim resa() As String
    For b = 1 To 50
    ReDim resa(0)
    For l = LBound(Quantities) To UBound(Quantities)
        If InStr(Quantities(l), achercher) > 0 Then
            If UBound(resa) = 0 Then
                ReDim resa(1 To 1)
            Else
                ReDim Preserve resa(1 To UBound(resa) + 1)
            End If
            resa(UBound(resa)) = Quantities(l)
        End If
    Next
    Next

    ta = Timer - sngChrono
    sngChrono = Timer
    For b = 1 To 50
    
    resb = Filter(Quantities, achercher, True)
    
    Next
    tb = Timer - sngChrono
    'MsgBox ta & ":" & UBound(resa) & Chr(10) & tb & ":" & UBound(resb)
    'MsgBox resa(1)
    'MsgBox resb(0)
End Sub
Function isAnAttr(attr As String, list() As String) As String
    listAttr = Split(Trim(attr), ",")
    If attr = "x" Then
    End If
    isAnAttr = ""
    For noAt = LBound(listAttr) To UBound(listAttr)
        If isAnAttr = "" Then jct = ""
        If isAnAttr <> "" Then jct = ","
        instring = Filter(list, listAttr(noAt), True)
        If (UBound(instring)) < 0 Then isAnAttr = isAnAttr & jct & listAttr(noAt)
    Next
End Function
Function delSpace(str As String) As String
    delSpace = Trim(str)
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Pattern = "[ ]+"
    If reg.test(delSpace) Then delSpace = reg.Replace(delSpace, "")
    reg.Pattern = "[ ]*,[ ]*"
    If reg.test(delSpace) Then delSpace = reg.Replace(delSpace, ",")
    reg.Pattern = "[ ]*\([ ]*"
    If reg.test(delSpace) Then delSpace = reg.Replace(delSpace, "(")
    'reg.Pattern = "[ ]*)[ ]*"
    'If reg.Test(delSpace) Then delSpace = reg.Replace(delSpace, ")")
End Function
Function testForme(liste As String, Nolig As Integer, ORIGINE As String) As Boolean
    testForme = True
    Dim newLogLine() As Variant
    list = Split(liste, "@")
    deuxieme = list(BEFORE2 - 1) & list(THIS2 - 1)
    If TestNotAlphaSup("" & list(BEFORE1 - 1)) Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "BEFORE1", list(BEFORE1 - 1), Nolig, "mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If TestNotAlpha("" & list(THIS1 - 1)) Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "THIS1", list(THIS1 - 1), Nolig, "mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If (list(BEFORE1 - 1) <> "" Or list(THIS1 - 1) <> "") And TestNotAlphaVir("" & list(ATTR1 - 1)) And deuxieme = "" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "ATTR1", list(ATTR1 - 1), Nolig, "attribut mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If (list(BEFORE1 - 1) <> "" Or list(THIS1 - 1) <> "") And TestNotAlphaVirOp("" & list(ATTR1 - 1)) And deuxieme <> "" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "ATTR1", list(ATTR1 - 1), Nolig, "attribut mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If list(BEFORE1 - 1) = "" And list(THIS1 - 1) = "" And TestNotAlpha("" & list(ATTR1 - 1)) Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "ATTR1", list(ATTR1 - 1), Nolig, "attribut mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If TestNotAlphaSup("" & list(BEFORE2 - 1)) Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "BEFORE2", list(BEFORE2 - 1), Nolig, "mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If TestNotAlpha("" & list(THIS2 - 1)) Then
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "THIS2", list(THIS2 - 1), Nolig, "mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
    If TestNotAlphaVirOp("" & list(ATTR2 - 1)) Then
'MsgBox "icic"
newLogLine = Array("", "ERREUR", ORIGINE, "Typographie", time(), "ATTR2", list(ATTR2 - 1), Nolig, "mal formé")
logNom = alimLog(logNom, newLogLine)
        testForme = False
    End If
End Function
Function extraitAttribut(att As String) As String
    extraitAttribut = att
    If InStr(att, "NOT(") = 1 Then
        extraitAttribut = Mid(att, 5)
        If InStr(extraitAttribut, ")") = Len(extraitAttribut) And Len(extraitAttribut) > 0 Then
            extraitAttribut = Left(extraitAttribut, Len(extraitAttribut) - 1)
        Else
            'mal formé
            extraitAttribut = att
        End If
    End If
    If InStr(att, "AND(") = 1 Then
        extraitAttribut = Mid(att, 5)
        If InStr(extraitAttribut, ")") = Len(extraitAttribut) And Len(extraitAttribut) > 0 Then
            extraitAttribut = Left(extraitAttribut, Len(extraitAttribut) - 1)
        Else
            'mal formé
            extraitAttribut = att
        End If
    End If
    If InStr(att, "OR(") = 1 Then
        extraitAttribut = Mid(att, 4)
        If InStr(extraitAttribut, ")") = Len(extraitAttribut) And Len(extraitAttribut) > 0 Then
            extraitAttribut = Left(extraitAttribut, Len(extraitAttribut) - 1)
        Else
            'mal formé
            extraitAttribut = att
        End If
    End If
    If InStr(att, "FIRST(") = 1 Then
        extraitAttribut = Mid(att, 7)
        If InStr(extraitAttribut, ")") = Len(extraitAttribut) And Len(extraitAttribut) > 0 Then
            extraitAttribut = Left(extraitAttribut, Len(extraitAttribut) - 1)
        Else
            'mal formé
            extraitAttribut = att
        End If
    End If
    If InStr(att, "LAST(") = 1 Then
        extraitAttribut = Mid(att, 6)
        If InStr(extraitAttribut, ")") = Len(extraitAttribut) And Len(extraitAttribut) > 0 Then
            extraitAttribut = Left(extraitAttribut, Len(extraitAttribut) - 1)
        Else
            'mal formé
            extraitAttribut = att
        End If
    End If
End Function
Function jonction(ligne As Integer, tableau() As Variant) As String
    If Trim(tableau(ligne, THIS1)) = "" Then jct1 = ""
    If Trim(tableau(ligne, THIS1)) <> "" Then jct1 = ">"
    If Trim(tableau(ligne, ATTR1)) = "" Then att1 = ""
    If Trim(tableau(ligne, ATTR1)) <> "" Then att1 = ":"
    jonction = Trim(tableau(ligne, BEFORE1)) & jct1 & Trim(tableau(ligne, THIS1)) & att1 & Trim(tableau(ligne, ATTR1))
    If Trim(tableau(ligne, BEFORE2)) <> "" Or Trim(tableau(ligne, THIS2)) <> "" Then
        j12 = "."
    Else
        j12 = ""
    End If
    If Trim(tableau(ligne, THIS2)) = "" Then jct2 = ""
    If Trim(tableau(ligne, THIS2)) <> "" Then jct2 = ">"
    If Trim(tableau(ligne, ATTR2)) = "" Then att2 = ""
    If Trim(tableau(ligne, ATTR2)) <> "" Then att2 = ":"
    jonction = jonction & j12 & Trim(tableau(ligne, BEFORE2)) & jct2 & Trim(tableau(ligne, THIS2)) & att2 & Trim(tableau(ligne, ATTR2))
    If Trim(tableau(ligne, BEFORE3)) <> "" Or Trim(tableau(ligne, THIS3)) <> "" Then
        j23 = "."
    Else
        j23 = ""
    End If
    If Trim(tableau(ligne, THIS3)) = "" Then jct3 = ""
    If Trim(tableau(ligne, THIS3)) <> "" Then jct3 = ">"
    If Trim(tableau(ligne, ATTR3)) = "" Then att3 = ""
    If Trim(tableau(ligne, ATTR3)) <> "" Then att3 = ":"
    jonction = jonction & j23 & Trim(tableau(ligne, BEFORE3)) & jct3 & Trim(tableau(ligne, THIS3)) & att3 & Trim(tableau(ligne, ATTR3))
    If Trim(tableau(ligne, BEFORE4)) <> "" Or Trim(tableau(ligne, THIS4)) <> "" Then
        j34 = "."
    Else
        j34 = ""
    End If
    If Trim(tableau(ligne, THIS4)) = "" Then jct4 = ""
    If Trim(tableau(ligne, THIS4)) <> "" Then jct4 = ">"
    If Trim(tableau(ligne, ATTR4)) = "" Then att4 = ""
    If Trim(tableau(ligne, ATTR4)) <> "" Then att4 = ":"
    jonction = jonction & j34 & Trim(tableau(ligne, BEFORE4)) & jct4 & Trim(tableau(ligne, THIS4)) & att4 & Trim(tableau(ligne, ATTR4))
    If Trim(tableau(ligne, BEFORE5)) <> "" Or Trim(tableau(ligne, THIS5)) <> "" Then
        j45 = "."
    Else
        j45 = ""
    End If
    If Trim(tableau(ligne, THIS5)) = "" Then jct5 = ""
    If Trim(tableau(ligne, THIS5)) <> "" Then jct5 = ">"
    If Trim(tableau(ligne, ATTR5)) = "" Then att5 = ""
    If Trim(tableau(ligne, ATTR5)) <> "" Then att5 = ":"
    jonction = jonction & j45 & Trim(tableau(ligne, BEFORE5)) & jct5 & Trim(tableau(ligne, THIS5)) & att5 & Trim(tableau(ligne, ATTR5))
    If Trim(tableau(ligne, BEFORE6)) <> "" Or Trim(tableau(ligne, THIS6)) <> "" Then
        j56 = "."
    Else
        j56 = ""
    End If
    If Trim(tableau(ligne, THIS6)) = "" Then jct6 = ""
    If Trim(tableau(ligne, THIS6)) <> "" Then jct6 = ">"
    If Trim(tableau(ligne, ATTR6)) = "" Then att6 = ""
    If Trim(tableau(ligne, ATTR6)) <> "" Then att6 = ":"
    jonction = jonction & j56 & Trim(tableau(ligne, BEFORE6)) & jct6 & Trim(tableau(ligne, THIS6)) & att6 & Trim(tableau(ligne, ATTR6))
    If Trim(tableau(ligne, BEFORE7)) <> "" Or Trim(tableau(ligne, THIS7)) <> "" Then
        j67 = "."
    Else
        j67 = ""
    End If
    If Trim(tableau(ligne, THIS7)) = "" Then jct7 = ""
    If Trim(tableau(ligne, THIS7)) <> "" Then jct7 = ">"
    If Trim(tableau(ligne, ATTR7)) = "" Then att7 = ""
    If Trim(tableau(ligne, ATTR7)) <> "" Then att7 = ":"
    jonction = jonction & j67 & Trim(tableau(ligne, BEFORE7)) & jct7 & Trim(tableau(ligne, THIS7)) & att7 & Trim(tableau(ligne, ATTR7))
    If Trim(tableau(ligne, BEFORE8)) <> "" Or Trim(tableau(ligne, THIS8)) <> "" Then
        j78 = "."
    Else
        j78 = ""
    End If
    If Trim(tableau(ligne, THIS8)) = "" Then jct8 = ""
    If Trim(tableau(ligne, THIS8)) <> "" Then jct8 = ">"
    If Trim(tableau(ligne, ATTR8)) = "" Then att8 = ""
    If Trim(tableau(ligne, ATTR8)) <> "" Then att8 = ":"
    jonction = jonction & j78 & Trim(tableau(ligne, BEFORE8)) & jct8 & Trim(tableau(ligne, THIS8)) & att8 & Trim(tableau(ligne, ATTR8))
    If Trim(tableau(ligne, BEFORE9)) <> "" Or Trim(tableau(ligne, THIS9)) <> "" Then
        j89 = "."
    Else
        j89 = ""
    End If
    If Trim(tableau(ligne, THIS9)) = "" Then jct9 = ""
    If Trim(tableau(ligne, THIS9)) <> "" Then jct9 = ">"
    If Trim(tableau(ligne, ATTR9)) = "" Then att9 = ""
    If Trim(tableau(ligne, ATTR9)) <> "" Then att9 = ":"
    jonction = jonction & j89 & Trim(tableau(ligne, BEFORE9)) & jct9 & Trim(tableau(ligne, THIS9)) & att9 & Trim(tableau(ligne, ATTR9))
End Function
Function compareListeListe(entites() As String, liste() As String) As Integer()
    Dim res() As Integer
    ReDim res(LBound(entites) To UBound(entites))
    For e = LBound(entites) To UBound(entites)
        res(e) = -1
        For l = LBound(liste) To UBound(liste)
            If liste(l) = entites(e) Then res(e) = l
        Next
    Next
    compareListeListe = res
End Function
Function propagationAttributs(entite As String, liste() As String, num As Integer) As String
    ' Propagation des attributs
    Dim joinSpl() As String
    Dim ind As Integer
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = ">"
    Dim jcta As String
'MsgBox entite
    Dim aj As String
        If entite Like "*.*" Then
            spl = Split(entite, ".")
            ReDim joinSpl(UBound(spl))
            ind = 0
            For NoF = LBound(spl) To UBound(spl)
                joinSpl(NoF) = spl(NoF)
            Next
            For Each facteur In spl
                ' on va chercher les attributs
                achercher = Split(facteur, ":")(0)
                Dim attr As String
                attr = ""
                For nol = LBound(liste) To UBound(liste)
                    If Split(liste(nol), ":")(0) = achercher Then
                        If UBound(Split(liste(nol), ":")) > 0 Then attr = Split(liste(nol), ":")(1)
                        Exit For
                    End If
                Next
                If UBound(Split(facteur, ":")) > 0 Then
                    ' cas ou il y a filtrage
                    ' RIEN A FAIRE ?
                Else
                    ' on pose les attributs
                    If attr <> "" Then joinSpl(ind) = facteur & ":" & attr '''all(Nol, 2)
                End If
                ind = ind + 1
            Next
            entite = Join(joinSpl, ".")
        Else
            ent = Split(entite, ":")(0)
            If UBound(Split(entite, ":")) > 0 Then
                ' Cas de surcharge des attributs
            Else
                ' Cas ou la chaine existe en entier dasn la liste
                achercher = entite
                Dim tag As Boolean
                tag = False
                For nol = LBound(liste) To UBound(liste)
                    If Split(liste(nol), ":")(0) = achercher Then
                        If UBound(Split(liste(nol), ":")) < 1 Then
                                jcta = ""
                                aj = ""
                        Else
                            jcta = ":"
                            aj = Trim(Split(liste(nol), ":")(1))
                        End If
                        entite = ent & jcta & aj
                        tag = True
                        Exit For
                    End If
                Next
'MsgBox tag & ":" & entite & ":::" & Join(liste, "|")
                If Not tag Then
                entR = StrReverse(Trim(ent))
                If InStr(entR, ">") > 0 Then
                    achercher = Left(ent, Len(ent) - InStr(entR, ">"))
                    For nol = LBound(liste) To UBound(liste)
                        If Split(liste(nol), ":")(0) = achercher Then
'MsgBox achercher & ":" & Split(liste(nol), ":")(0)
                            If UBound(Split(liste(nol), ":")) < 1 Then
                                jcta = ""
                                aj = ""
                            Else
                                jcta = ":"
                                aj = Trim(Split(liste(nol), ":")(1))
                            End If
                            entite = ent & jcta & aj
'MsgBox aj & ":" & Split(liste(nol), ":")(0)
                            Exit For
                        End If
                    Next
                End If
                End If
            End If
        End If
'MsgBox "res=" & entite & "::" & Join(liste, "|")
propagationAttributs = entite
End Function

Function isAnAttribut(attr As String, ATTRIB() As String, num As Integer, dimension As String, ORIGINE As String, colonne As String) As Boolean
    ' renvoie False si un des attributs de la liste n'en est pas un
    isAnAttribut = True
    Dim attribs() As String
    ReDim attribs(0)
    splAttr = Split(attr, ",")
    Dim I As Integer
    aa = 0
    Dim tag As Boolean
    For a = LBound(splAttr) To UBound(splAttr)
        tag = False
        For I = LBound(ATTRIB) To UBound(ATTRIB)
            If splAttr(a) = ATTRIB(I) Then
'If num = 100 Then MsgBox "isAnAttribut>>" & splAttr(a) & "<<<>>>" & ATTRIB(i)
                tag = True
                Exit For
            End If
        Next
        If Not tag Then
            aa = aa + 1
'If num = 100 Then MsgBox "isAnAttribut>>" & splAttr(a)
            If UBound(attribs) = 0 Then
                ReDim attribs(1 To aa)
            Else
                ReDim Preserve attribs(1 To aa)
            End If
            attribs(aa) = splAttr(a)
        End If
    Next
'If num = 100 Then MsgBox attr & "<<<>>>" & tag
'If num = 145 Then MsgBox "isAnAttribut>>" & UBound(logNom, 1) & ":" & UBound(logNom, 2) & "::" & logNom(8, UBound(logNom, 2))
    If UBound(attribs) > 0 Then
        isAnAttribut = False
'''MsgBox num & ":"
        If logNom(8, UBound(logNom, 2)) <> ("" & num) Or logNom(9, UBound(logNom, 2)) <> "attribut non défini" Then
            Dim newLogLine() As Variant
'MsgBox colonne & ":" & dimension & ":" & num & ":attr=" & attr
newLogLine = Array("", "ALERTE", ORIGINE, "Attribut", time(), colonne, Join(attribs, ";"), num, "attribut non défini")
logNom = alimLog(logNom, newLogLine)
        End If
    End If
'If num = 145 Then MsgBox "isAnAttribut>>" & UBound(logNom, 1) & ":" & UBound(logNom, 2) & "::" & logNom(8, UBound(logNom, 2))
End Function





''' Construction dela nomenclature étendue à partir dela nomenclature générique
Sub SetNomenclature()
    '''On Error GoTo errorHandler
    If Not openFileIfNot("MODELE") Then Exit Sub
    ReDim ATTRIBUTS(0)
    Dim FL0QU As Worksheet
    Dim FLCTL As Worksheet
    Dim sngChrono As Single
    sngChrono = Timer
    'Dim FL0QU As Worksheet
    Dim FLCNO As Worksheet
    Set FLCTL = Worksheets(g_CONTROL)
    Dim nameSheetEnCours As String
Application.StatusBar = "CONSTRUCTION DES QUANTITES"
    '''If ActiveSheet.NAME = "CONTROL" Then
        'numnom = CInt(FLCTL.Cells(1, 5).VALUE)
        'nameSheetEnCours = Trim(FLCTL.Cells(1 + numnom, 5).VALUE)
    nameSheetEnCours = ActiveSheet.cbNomenclature.VALUE
    '''Else
        '''nameSheetEnCours = ActiveSheet.NAME
    '''End If
'MsgBox nameSheetEnCours & Chr(10) & WsExistInWB(nameSheetEnCours, g_WB_Modele)
    Set FL0QU = g_WB_Modele.Worksheets(nameSheetEnCours)
    Set FLCNO = g_WB_Modele.Worksheets("CRNOM")
    Set FLNOM = g_WB_Modele.Worksheets("NOMENCLATURE")
    Set FLAREA = g_WB_Modele.Worksheets("AREA")
    Set FLSCENARIO = g_WB_Modele.Worksheets("SCENARIO")
    FLNOM.Cells.Clear
    resInitGlobales = initGlobales()
    Dim LigDeb0No As Integer
    derlig = Split(FLCNO.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLCNO.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    derligsheet = FLCNO.Range("A" & FLCNO.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCNO.Columns.Count).End(xlToLeft).Column
    FLCNO.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    With FLCNO.Range("a" & (LigDebCNo - 1) & ":" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Dim newLogLine() As Variant
    newLogLine = Array("NOMENCLATURE", "DEBUT", FL0QU.NAME, "GENERATION", time())
    logNom = alimLog(logNom, newLogLine)

    derlig = FL0QU.Cells.SpecialCells(xlCellTypeLastCell).Row
    For Nolig = 1 To derlig
        If FL0QU.Cells(Nolig, 1) = "ACTION" Then
            LigDeb0No = Nolig + 1
            Exit For
        End If
    Next
    
    'Alimentation des tableaux
    Dim Cla() As Variant

    'DerLigSheet = FL0NO.Range("A" & Rows.Count).End(xlUp).Row
    derligsheet = getDerLig(FL0QU)
    'DerLigSheet = FL0QU.Cells.SpecialCells(xlCellTypeLastCell).Row
    DerColSheet = getDerCol(FL0QU)
    'DerColSheet = Cells(LigDeb0No - 1, Columns.Count).End(xlToLeft).Column

    With FL0QU.Range("a" & LigDeb0No & ":" & "AF" & derligsheet)
        ReDim Cla(1 To (derligsheet - LigDeb0No), 1 To DerColSheet)
        Cla = .VALUE
    End With
    ' boucle principale
    Dim eten() As String
    Dim lig As Integer
    Dim entite As String
    Dim error As Boolean
    Dim listRes() As String
    ReDim listRes(0)
    Dim listShort() As String
    ReDim listShort(0)
    Dim listName() As String
    ReDim listName(0)
    Dim listSyn() As String
    ReDim listSyn(0)
    Dim listID() As String
    ReDim listID(0)
    Dim listNb() As String
    Dim listResNew() As String
    ReDim listResNew(0)
    Dim ind As Integer
    ind = 0
    Dim numlig() As String
    ReDim numlig(0)
    Dim numLigNew() As String
    ReDim numLigNew(0)
    Dim nextLigne As Integer
    Dim nbAjout As Integer
    Dim listSurcharge() As Boolean
    Dim ligSurcharge As String
    ligSurcharge = ""
    Dim cibSurcharge As String
    cibSurcharge = ""
    Dim ligRetrait As String
    ligRetrait = ""
    Dim cibRetrait As String
    cibRetrait = ""
    Dim nbSurcharge As Integer
    nbSurcharge = 0
    Dim listMinus() As Boolean
    Dim nbRealMinus As Integer
    Dim posAttr As Integer
    'Dim attributs() As String
    Dim vecteur As String
    ReDim listNb(1 To UBound(Cla, 1), 1 To 1)
    Dim nbRetraitsEnTout As Integer
    nbRetraitsEnTout = 0
    Dim nbLigOK As Integer
    nbLigOK = 0
    Dim nbAttOK As Integer
    nbAttOK = 0
    Dim compteur As Integer
    Dim errornb As Integer
    errornb = 0
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        listNb(Nolig, 1) = ""
compteur = 100 * Nolig / UBound(Cla, 1)
Application.StatusBar = "CONSTRUCTION DE LA NOMENCLATURE : GENERATION : " & Nolig & " / " & UBound(Cla, 1) & " " & compteur & " %"
        If Left(Cla(Nolig, 1), 1) <> "#" Then
            vecteur = nettoyage(Cla(Nolig, ACTION)) & "@" & nettoyage(Cla(Nolig, ID)) & "@" & nettoyage(Cla(Nolig, BEFORE1)) _
                & "@" & nettoyage(Cla(Nolig, THIS1)) & "@" & nettoyage(Cla(Nolig, ATTR1)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE2)) & "@" & nettoyage(Cla(Nolig, THIS2)) & "@" & nettoyage(Cla(Nolig, ATTR2)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE3)) & "@" & nettoyage(Cla(Nolig, THIS3)) & "@" & nettoyage(Cla(Nolig, ATTR3)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE4)) & "@" & nettoyage(Cla(Nolig, THIS4)) & "@" & nettoyage(Cla(Nolig, ATTR4)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE5)) & "@" & nettoyage(Cla(Nolig, THIS5)) & "@" & nettoyage(Cla(Nolig, ATTR5)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE6)) & "@" & nettoyage(Cla(Nolig, THIS6)) & "@" & nettoyage(Cla(Nolig, ATTR6)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE7)) & "@" & nettoyage(Cla(Nolig, THIS7)) & "@" & nettoyage(Cla(Nolig, ATTR7)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE8)) & "@" & nettoyage(Cla(Nolig, THIS8)) & "@" & nettoyage(Cla(Nolig, ATTR8)) _
                & "@" & nettoyage(Cla(Nolig, BEFORE9)) & "@" & nettoyage(Cla(Nolig, THIS9)) & "@" & nettoyage(Cla(Nolig, ATTR9))
            resTestForme = testForme(vecteur, LigDeb0No + Nolig - 1, FL0QU.NAME)
'If (LigDeb0No + Nolig - 1) = 64 Then MsgBox extraitAttribut(nettoyage(Cla(Nolig, ATTR1))) & Chr(10) & vecteur
            If Not resTestForme Then
                listNb(Nolig, 1) = "0"
                GoTo continueLoop
            End If
            listNb(Nolig, 1) = "1"
            If Cla(Nolig, BEFORE1) = "" And Cla(Nolig, THIS1) = "" And Cla(Nolig, ATTR1) <> "" Then
                posAttr = posAttr + 1
                If UBound(ATTRIBUTS) = 0 Then
                    ReDim ATTRIBUTS(1 To 1)
                Else
                    ReDim Preserve ATTRIBUTS(1 To posAttr)
                End If
                ATTRIBUTS(posAttr) = Trim(Cla(Nolig, ATTR1))
'If (LigDeb0No + Nolig - 1) = 64 Then MsgBox ATTRIBUTS(posAttr)
            Else
            error = False
            lig = Nolig
            entite = jonction(lig, Cla)
'If (LigDeb0No + Nolig - 1) = 64 Then MsgBox entite
            If entite = "" Then
                listNb(Nolig, 1) = ""
                GoTo continueLoop
            End If
            ' résolution des raccourcis
            If UBound(listRes) <> 0 Then
                entite = remplaceDoubleSup(entite, listRes, FL0QU.NAME, LigDeb0No + Nolig - 1)
            End If
            If entite Like "*>>*" Then
                splEntite = Split(entite, ".")
                where = ""
                For I = LBound(splEntite) To UBound(splEntite)
                    If splEntite(I) Like "*>>*" Then where = where & "," & "BEFORE" & (I + 1)
                Next
                where = Mid(where, 2)
errornb = errornb + 1
newLogLine = Array("", "ERREUR", FL0QU.NAME, "Raccourci", time(), where, entite, LigDeb0No + Nolig - 1, "raccourci non résolu")
logNom = alimLog(logNom, newLogLine)
                error = True
            End If
            If error Then
                listNb(Nolig, 1) = "0"
                GoTo continueLoop
            End If
            ' test de prédéfinition de l'antécédent
            If UBound(Split(entite, ".")) = 0 And Not getPre(entite, listRes) Then
newLogLine = Array("", "ALERTE", FL0QU.NAME, "Antécédent", time(), "BEFORE1", entite, LigDeb0No + Nolig - 1, "antécédent non défini")
logNom = alimLog(logNom, newLogLine)
            End If
            ''' true ou false ???
'If (LigDeb0No + Nolig - 1) = 64 Then MsgBox "debut = setqua=" & entite
            'attention listRes contient des entités complexes (avec un point) !!!
            eten = getExtended("nom", False, entite, listRes, LigDeb0No + Nolig - 1, "NOMENCLATURE", FL0QU.NAME, "L")
'If (LigDeb0No + Nolig - 1) = 64 Then MsgBox "fin = setqua=" & Join(eten, Chr(10))
            listNb(Nolig, 1) = "" & UBound(eten)
            If UBound(eten) > 0 Then
                ReDim listSurcharge(1 To (UBound(eten)))
                nextLigne = 0
                nbAjout = 0
                oldUb = UBound(listRes)
                ' Surcharge
                If Trim(Cla(Nolig, ACTION)) <> "minus" Then
                    For ne = LBound(eten) To UBound(eten)
                        nextLigne = nextLigne + 1
                        listSurcharge(nextLigne) = False
                        nbAjout = nbAjout + 1
                        If UBound(numlig) > 0 Then
                            listSur = ""
                            For I = LBound(numlig) To UBound(numlig)
                                ' ????
                                splEten = Split(eten(ne), ".")
                                etenSansAtt = Split(eten(ne), ":")(0)
                                resSansAtt = Split(listRes(I), ":")(0)
                                If UBound(splEten) > 0 Then
                                    etenATester = eten(ne)
                                    listResATester = listRes(I)
                                Else
                                    etenATester = etenSansAtt
                                    listResATester = resSansAtt
                                End If
                                If etenATester = listResATester Then
                                    ' surcharge ici
'MsgBox listRes(i) & Chr(10) & eten(ne)
                                    listSurcharge(nextLigne) = True
                                    nbAjout = nbAjout - 1
                                    nbSurcharge = nbSurcharge + 1
                                    If ligSurcharge <> "" Then lastLigSurcharge = Split(ligSurcharge, ";")(UBound(Split(ligSurcharge, ";")))
                                    If ligSurcharge = "" Then lastLigSurcharge = ""
                                    If lastLigSurcharge <> ("" & (LigDeb0No + Nolig - 1)) Then
                                        ligSurcharge = ligSurcharge & ";" & (LigDeb0No + Nolig - 1)
' If (LigDeb0No + Nolig - 1) = 115 Then MsgBox ">>>>>>>" & ligSurcharge & "=>" & lastLigSurcharge & "<==>" & Split(ligSurcharge, ";")(UBound(Split(ligSurcharge, ";"))) & "<"
                                    End If
'If (LigDeb0No + Nolig - 1) = 115 Then MsgBox ligSurcharge & "=" & lastLigSurcharge
                                    If listSur <> "" Then listSur = listSur & "," & numlig(I)
                                    If listSur = "" Then listSur = "" & numlig(I)
                                    numlig(I) = numlig(I) & "," & (LigDeb0No + Nolig - 1)
                                    lastResI = listRes(I)
                                    listRes(I) = eten(ne)
                                    'listShort(i) = Split(Split(Split(eten(ne), ".")(UBound(Split(eten(ne), "."))), ":")(0), ">")(UBound(Split(Split(Split(eten(ne), ".")(UBound(Split(eten(ne), "."))), ":")(0), ">")))
                                    listName(I) = Trim(Cla(Nolig, NAME))
                                    listID(I) = Trim(Cla(Nolig, ID))
                                    listSyn(I) = Trim(Cla(Nolig, SYNONYMOUS))
'If Trim(Cla(Nolig, SYNONYMOUS)) <> "" Then MsgBox Trim(Cla(Nolig, SYNONYMOUS))
                                    ' Il faut répercuter un éventuel ajout d'attribut sur les entités qui le contiennent
                                    'If lastResI <> listRes(i) And (Not InStr(listRes(i), ".") > 0) Then
                                        'MsgBox lastResI & Chr(10) & eten(ne) & Chr(10) & i & "<" & UBound(listRes)
                                        'For t = i + 1 To UBound(listRes)
'If InStr(">" & listRes(t) & ">", ">" & lastResI & ">") > 0 Or InStr(">" & listRes(t) & ".", ">" & lastResI & ".") > 0 Then
                                                'listRes(t) = Replace(listRes(t), lastResI, listRes(i))
                                            'End If
'If InStr("." & listRes(t) & ">", "." & lastResI & ">") > 0 Or InStr("." & listRes(t) & ".", "." & lastResI & ".") > 0 Then
                                                'listRes(t) = Replace(listRes(t), lastResI, listRes(i))
                                            'End If
                                        'Next
                                    'End If
                                    Exit For
                                End If
                            Next
                            If listSur <> "" Then
                                If lastLigSurcharge <> ("" & (LigDeb0No + Nolig - 1)) Then
                                'lastListSur = Split(listSur, ";")(0)
                                    cibSurcharge = cibSurcharge & ";" & listSur
                                End If
                            End If
                        End If
                    Next
                
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":::" & listSur & "<<<"
                End If
                If Trim(Cla(Nolig, ACTION)) = "minus" Then
'If (LigDeb0No + Nolig - 1) = 122 Then MsgBox UBound(eten) & ":" & vecteur
                    ReDim listMinus(LBound(numlig) To UBound(numlig))
                    nextLigne = 0
                    nbMinus = 0
                    nbRealMinus = 0
                    listCibEnCours = ""
                    For I = LBound(numlig) To UBound(numlig)
                        listMinus(I) = False
                    Next
                    For ne = LBound(eten) To UBound(eten)
                        nextLigne = nextLigne + 1
                        nbMinus = nbMinus + 1
                        listCib = ""
                        For I = LBound(numlig) To UBound(numlig)
                            '''listMinus(i) = False
                            If eten(ne) = listRes(I) Then
                                listMinus(I) = True
                                Dim lastRetrait As String
                                If ligRetrait <> "" Then
                                    lastRetrait = Split(ligRetrait, ";")(UBound(Split(ligRetrait, ";")))
                                    If lastRetrait <> ("" & (LigDeb0No + Nolig - 1)) Then ligRetrait = ligRetrait & ";" & (LigDeb0No + Nolig - 1)
'If (LigDeb0No + Nolig - 1) = 122 Then MsgBox lastRetrait & ">>>" & ligRetrait
                                End If
                                If ligRetrait = "" Then ligRetrait = "" & (LigDeb0No + Nolig - 1)
                                '''If listCibEnCours <> "" Then listCibEnCours = listCibEnCours & "," & numLig(i)
                                If listCibEnCours <> "" Then
                                    lastRetrait = Split(listCibEnCours, ";")(UBound(Split(listCibEnCours, ";")))
                                    If lastRetrait <> ("" & numlig(I)) Then listCibEnCours = listCibEnCours & "," & numlig(I)
                                End If
                                If listCibEnCours = "" Then listCibEnCours = "" & numlig(I)
                                nbRealMinus = nbRealMinus + 1
                                nbRetraitsEnTout = nbRetraitsEnTout + 1
'If (LigDeb0No + Nolig - 1) = 122 Then MsgBox UBound(eten) & ":" & eten(ne) & "::" & numLig(i)
                                'Exit For
                            End If
                        Next
                    Next
'If (LigDeb0No + Nolig - 1) = 122 Then MsgBox UBound(eten) & ":" & nbRealMinus
                    If nbRealMinus > 0 Then
                        If cibRetrait <> "" Then cibRetrait = cibRetrait & ";" & listCibEnCours
                        If cibRetrait = "" Then cibRetrait = listCibEnCours
                        ReDim listResNew(1 To (UBound(listRes) - nbRealMinus))
                        ReDim numLigNew(1 To (UBound(numlig) - nbRealMinus))
                        ReDim listShortNew(1 To (UBound(listShort) - nbRealMinus))
                        ReDim listNameNew(1 To (UBound(listName) - nbRealMinus))
                        ReDim listSynNew(1 To (UBound(listSyn) - nbRealMinus))
                        ReDim listIDNew(1 To (UBound(listID) - nbRealMinus))
                        Dim j As Integer
                        j = 0
                        For I = LBound(listRes) To UBound(listRes)
                            If Not listMinus(I) Then
                                j = j + 1 ' If j <=UBound(listResNew)
                                If j <= UBound(listResNew) Then listResNew(j) = listRes(I)
                                If j <= UBound(listResNew) Then numLigNew(j) = numlig(I)
                                If j <= UBound(listResNew) Then listShortNew(j) = listShort(I)
                                If j <= UBound(listResNew) Then listNameNew(j) = listName(I)
                                If j <= UBound(listResNew) Then listSynNew(j) = listSyn(I)
                                If j <= UBound(listResNew) Then listIDNew(j) = listID(I)
                            Else
'If (LigDeb0No + Nolig - 1) = 122 Then MsgBox UBound(eten) & ":" & j & ":" & listRes(i)
                            End If
                        Next
                        ReDim listRes(1 To UBound(listResNew))
                        ReDim numlig(1 To UBound(numLigNew))
                        ReDim listShort(1 To UBound(listShortNew))
                        ReDim listName(1 To UBound(listNameNew))
                        ReDim listSyn(1 To UBound(listSynNew))
                        ReDim listID(1 To UBound(listIDNew))
                        For I = LBound(listResNew) To UBound(listResNew)
                            listRes(I) = listResNew(I)
                            numlig(I) = numLigNew(I)
                            listShort(I) = listShortNew(I)
                            listName(I) = listNameNew(I)
                            listSyn(I) = listSynNew(I)
                            listID(I) = listIDNew(I)
                        Next
                        listNb(Nolig, 1) = "-" & nbRealMinus
                    Else
newLogLine = Array("", "ALERTE", FL0QU.NAME, "Retrait", time(), "", entite, (LigDeb0No + Nolig - 1), "Retrait non résolu")
logNom = alimLog(logNom, newLogLine)
                    listNb(Nolig, 1) = "0"
                    End If
                Else
'If (LigDeb0No + Nolig - 1) = 115 Then MsgBox (LigDeb0No + Nolig - 1) & "==" & nbAjout & "==" & UBound(eten) & "===" & UBound(listRes)
                    If ind = 0 Then
                        ReDim listRes(1 To (UBound(listRes) + nbAjout))
                        ReDim numlig(1 To (UBound(numlig) + nbAjout))
                        ReDim listShort(1 To (UBound(listShort) + nbAjout))
                        ReDim listName(1 To (UBound(listName) + nbAjout))
                        ReDim listSyn(1 To (UBound(listSyn) + nbAjout))
                        ReDim listID(1 To (UBound(listID) + nbAjout))
                    Else
                        ReDim Preserve listRes(1 To (UBound(listRes) + nbAjout))
                        ReDim Preserve numlig(1 To (UBound(numlig) + nbAjout))
                        ReDim Preserve listShort(1 To (UBound(listShort) + nbAjout))
                        ReDim Preserve listName(1 To (UBound(listName) + nbAjout))
                        ReDim Preserve listSyn(1 To (UBound(listSyn) + nbAjout))
                        ReDim Preserve listID(1 To (UBound(listID) + nbAjout))
                    End If
'If (LigDeb0No + Nolig - 1) = 115 Then MsgBox (LigDeb0No + Nolig - 1) & "==" & UBound(eten) & "===" & UBound(listRes)
                    ind = ind + nbAjout
'If (LigDeb0No + Nolig - 1) = 117 Then MsgBox (LigDeb0No + Nolig - 1) & ">>>>>>nb eten=" & UBound(eten) & ":nbAjout=" & nbAjout & ":" & listRes(UBound(listRes)) & "<<<" & ind
                    nextLigne = 0
'If (LigDeb0No + Nolig - 1) = 117 Then MsgBox ">>>>" & Join(listRes, "|") & "<<<<"
                    Dim nextAjout As Integer
                    nextAjout = 0
                    For ne = LBound(eten) To UBound(eten)
                        nextLigne = nextLigne + 1
                        If Not listSurcharge(nextLigne) Then
                            nextAjout = nextAjout + 1
                            listRes(oldUb + nextAjout) = eten(ne)
                            numlig(oldUb + nextAjout) = "" & (LigDeb0No + Nolig - 1)
                            listShort(oldUb + nextAjout) = Split(Split(Split(eten(ne), ".")(UBound(Split(eten(ne), "."))), ":")(0), ">")(UBound(Split(Split(Split(eten(ne), ".")(UBound(Split(eten(ne), "."))), ":")(0), ">")))
                            listName(oldUb + nextAjout) = Trim(Cla(Nolig, NAME))
                            listSyn(oldUb + nextAjout) = Replace(Trim(Cla(Nolig, SYNONYMOUS)), " ,", ",")
                            listID(oldUb + nextAjout) = Trim(Cla(Nolig, ID))
                        End If
                    Next
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & UBound(listRes) & ":" & listRes(UBound(listRes) - nbAjout + 1)
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":" & listRes(UBound(listRes) - nbAjout + 2)
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":" & listRes(UBound(listRes) - nbAjout + 3)
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":" & listRes(UBound(listRes) - nbAjout + 4)
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":" & listRes(UBound(listRes) - nbAjout + 5)
'If (LigDeb0No + Nolig - 1) = 118 Then MsgBox UBound(eten) & ":" & nbAjout & ":" & listRes(UBound(listRes) - nbAjout + 6)
                End If
            End If
            End If
            nbLigOK = nbLigOK + 1
        End If
continueLoop:
    Next
    If ligSurcharge <> "" Then ligSurcharge = Mid(ligSurcharge, 2)
    'If ligRetrait <> "" Then ligRetrait = Mid(ligRetrait, 2)
    If cibSurcharge <> "" Then cibSurcharge = Mid(cibSurcharge, 2)
    'If cibRetrait <> "" Then cibRetrait = Mid(cibRetrait, 2)
    Dim aecrire() As String
    ReDim aecrire(1 To UBound(listRes), 1 To 7)
    For Nolig = 1 To UBound(listRes)
        aecrire(Nolig, 1) = FL0QU.NAME
        aecrire(Nolig, 2) = numlig(Nolig)
        aecrire(Nolig, 3) = listRes(Nolig)
        aecrire(Nolig, 4) = listShort(Nolig)
        aecrire(Nolig, 5) = listName(Nolig)
        aecrire(Nolig, 6) = listID(Nolig)
        aecrire(Nolig, 7) = listSyn(Nolig)
    Next
    ' Ecriture des résultats
    Dim scs As String
    If nbSurcharge > 1 Then scs = "s"
    If nbSurcharge < 2 Then scs = ""
    Dim rts As String
    If nbRetraitsEnTout > 1 Then rts = "s"
    If nbRetraitsEnTout < 2 Then rts = ""
    Dim lgs As String
    If nbLigOK > 1 Then lgs = "s"
    If nbLigOK < 2 Then lgs = ""
    Dim ets As String
    If UBound(listRes) > 1 Then ets = "s"
    If UBound(listRes) < 2 Then ets = ""
    Dim ats As String
    If UBound(ATTRIBUTS) > 1 Then ats = "s"
    If UBound(ATTRIBUTS) < 2 Then ats = ""
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", cibSurcharge, ligSurcharge, nbSurcharge & " surcharge" & scs)
logNom = alimLog(logNom, newLogLine)
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", cibRetrait, ligRetrait, nbRetraitsEnTout & " retrait" & rts)
logNom = alimLog(logNom, newLogLine)
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", "", nbLigOK & " ligne" & lgs, UBound(ATTRIBUTS) & " attribut" & ats & " et " & UBound(listRes) & " entité" & ets & " générée" & ets)
logNom = alimLog(logNom, newLogLine)
    Dim entete() As String
    ReDim entete(1 To 1, 1 To 8)
    entete(1, 1) = "Feuille"
    entete(1, 2) = "Ligne"
    entete(1, 3) = "ENTITE"
    entete(1, 4) = "SHORTNAME"
    entete(1, 5) = "NAME"
    entete(1, 6) = "ID"
    entete(1, 7) = "SYNONYMOUS"
    entete(1, 8) = "ATTRIBUT"
    FLNOM.Range("A1:H1").VALUE = entete
    If UBound(ATTRIBUTS) > 0 Then
        Dim attribaec() As String
        ReDim attribaec(1 To UBound(ATTRIBUTS), 1 To 1)
        For I = LBound(ATTRIBUTS) To UBound(ATTRIBUTS)
            attribaec(I, 1) = ATTRIBUTS(I)
        'MsgBox attribaec(i, 1)
        Next
    End If
'MsgBox Join(ATTRIBUTS, Chr(10))
    FLNOM.Range("A2:G" & (UBound(listRes) + 1)).VALUE = aecrire
    If UBound(ATTRIBUTS) > 0 Then FLNOM.Range("H2:H" & (UBound(ATTRIBUTS) + 1)).VALUE = attribaec
    FL0QU.Range("AD" & LigDeb0No & ":AD" & (LigDeb0No + UBound(listNb, 1) - 1)).VALUE = listNb
    sngChrono = Timer - sngChrono
    newLogLine = Array("NOMENCLATURE", "FIN", FL0QU.NAME, "GENERATION", time(), "", "", "", "Durée = " & (Int(1000 * sngChrono) / 1000) & " s")
    logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCNO.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    'FLCNO.Range("A1:K" & UBound(logNom, 1)).Borders(xlEdgeTop).Weight = xlSolid
    FLCNO.Range("A1:K" & UBound(logNom, 1)).Cells.Borders.LineStyle = xlContinuous
    resColoriage = coloriage(LigDeb0No, FL0QU, FLCNO)
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C).VALUE = nbLigOK & ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 1).VALUE = UBound(listRes) & ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 2).VALUE = resColoriage & ""
    FLCTL.Cells(g_CONTROL_NOMENCLATURE_GEN_CR_L, g_CONTROL_NOMENCLATURE_GEN_CR_C + 3).VALUE = Round((Int(1000 * sngChrono) / 1000)) & ""
Application.StatusBar = False
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setNomenclature"): Exit Sub
End Sub
Function lectureTime(feuille As String) As String()
    Set FLTIME = g_WB_Modele.Worksheets(feuille)
    Dim res() As String
    Dim areaList() As String
    Dim derlig As Integer
    Dim dercol As Integer
    ReDim res(0)
    Dim Cla() As Variant
    derlig = FLTIME.Range("A" & FLTIME.Rows.Count).End(xlUp).Row
    dercol = FLTIME.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLTIME.Range("a1" & ":" & DecAlph(dercol) & derlig)
        ReDim Cla(1 To derlig, 1 To dercol)
        Cla = .VALUE
    End With
    Dim ResEnCours As String
    ResEnCours = ""
    Dim lastcat As String
    lastcat = ""
    Dim cat As String
    cat = ""
    Dim typ As String
    typ = ""
    Dim list() As String
    ReDim list(0)
    Dim ok As Boolean
    ok = False
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        ResEnCours = ""
        ReDim list(0)
        If Left(Cla(Nolig, 1), 1) <> "#" Then
            If Nolig > 1 Then
                typ = Trim(Cla(Nolig, 1))
                lastcat = ""
                For nol = 2 To dercol
                    cat = Trim(Cla(Nolig, nol))
                    DAT = Trim(Cla(1, nol))
                    If cat <> "" Then
                        If UBound(list) > 0 Then
                            ok = False
                            For I = LBound(list) To UBound(list)
                                splList0 = Split(list(I), "@")(0)
                                If splList0 = (typ & "$" & cat) Then
                                    jct = ","
                                    ResEnCours = list(I) & jct & DAT
                                    list(I) = ResEnCours
                                    ok = True
                                    Exit For
                                End If
                                lastcat = cat
                            Next
                            If Not ok Then
                                taille = UBound(list) + 1
                                ReDim Preserve list(1 To taille)
                                ResEnCours = typ & "$" & cat & "@" & DAT
                                list(taille) = ResEnCours
                                
                            End If
                        Else
                            ReDim list(1 To 1)
                            ResEnCours = typ & "$" & cat & "@" & DAT
                            list(1) = ResEnCours
                        End If
                    End If
                Next
                If UBound(list) > 0 Then
                    oldtaille = UBound(res)
                    taille = UBound(res) + UBound(list)
                    If UBound(res) > 0 Then ReDim Preserve res(1 To taille)
                    If UBound(res) = 0 Then ReDim res(1 To taille)
                    For j = 1 To UBound(list)
                        res(oldtaille + j) = list(j)
                    Next
                End If
            Else
                ResEnCours = "$@"
                For nol = 3 To dercol
                    If nol = 3 Then
                        jct = ""
                    Else
                        jct = ","
                    End If
                    ResEnCours = ResEnCours & jct & Trim(Cla(1, nol))
                Next
                ReDim res(1 To 1)
                res(1) = ResEnCours
            End If
        End If
    Next
    lectureTime = res
End Function
Function isATime(tim As String, list() As String) As Boolean
    isATime = False
    If tim <> "" Then
        Dim deb As String
        Dim debt As String
        Dim sec As String
        Dim listDates As String
        debt = Split(tim, "$")(0)
        If IsNumeric(tim) Then
            listDates = Split(list(LBound(list)), "@")(1)
            spllistdates = Split(listDates, ",")
            For d = LBound(spllistdates) To UBound(spllistdates)
                If spllistdates(d) = "" & tim Then
                    isATime = True
                    Exit For
                End If
            Next
        Else
            For I = LBound(list) To UBound(list)
                deb = Split(list(I), "@")(0)
                sec = "$" & Split(deb, "$")(1)
                If InStr(tim, ">") > 0 Then
                    timav = Split(tim, ">")(0)
                    If UBound(Split(tim, ">")) > 0 Then
                        timap = Split(tim, ">")(1)
                    Else
                        timpa = ""
                    End If
                    If timap <> "" And LCase(timap) <> "last" And LCase(timap) <> "first" Then
                        isATime = False
                        Exit For
                    End If
                    If timav = deb Then
                        isATime = True
                        Exit For
                    End If
                    If debt = "" And sec = timav Then
                        isATime = True
                        Exit For
                    End If
                Else
                    timav = Split(tim, ".")(0)
                    If UBound(Split(tim, ".")) > 0 Then
                        timap = Split(tim, ".")(1)
                    Else
                        timpa = ""
                    End If
                    If timap <> "" And timap <> "last" And timap <> "first" Then
                        isATime = False
                        Exit For
                    End If
                    If timav = deb Then
                        isATime = True
                        Exit For
                    End If
                    If debt = "" And sec = timav Then
                        isATime = True
                        Exit For
                    End If
                End If
            Next
        End If
    End If
End Function

Function getPerimeter(perT As String, perI As String, perO As String, listN() As String, listA() As String, listS() As String, num As Integer, FL0 As Worksheet, listN2() As String, listA2() As String, listS2() As String, listQ2() As String) As String
' extension et test d'un périmètre
'If num = 12 And perT Like "*THIS2*" Then MsgBox perT & Chr(10) & perI
    ' perT : périmètre à tranformer
    ' perI : périmètre en input (ne sert que pour les attributs???)
    ' perO : périmètre étendu
    getPerimeter = perO
    Dim spldimT() As String
    spldimT = Split(perT, "]")
    splDimI = Split(perI, "]")
    splDimO = Split(perO, "]")
    Dim aTraiter As String
    Dim outThis1 As String
    Dim outThis2 As String
    Dim outThis3 As String
    Dim outThis As String
    Dim res As String
    Dim eten() As String
    ' ajout des dimensions absentes ou []
'If num = 22 Then MsgBox "getPerimeter:" & perT & Chr(10) & perI & Chr(10) & perO
    'entite = remplaceDoubleSup(entite, NOMENCLA, FL0QU.NAME, 0)
    If UBound(spldimT) < UBound(splDimO) Then
        Dim qua As String
        qua = spldimT(UBound(spldimT))
        For I = UBound(spldimT) To UBound(splDimO) - 1
'If num = 16 Then MsgBox LBound(spldimT) & " to " & i
            ReDim Preserve spldimT(0 To I)
            spldimT(I) = splDimO(I)
        Next
        ReDim Preserve spldimT(0 To UBound(splDimO))
        spldimT(I) = qua
'MsgBox perT & "===" & splDimO(i) & ":::" & i
    
    End If
'If num = 16 Then MsgBox Join(spldimT, "]") & "::" & LBound(spldimT) & "::" & UBound(spldimT) & "::" & UBound(spldimO)
    For I = LBound(spldimT) To UBound(spldimT) - 1
        If spldimT(I) = "[" Then
'If num = 16 Then MsgBox spldimT(i) & "===" & splDimO(i)
            spldimT(I) = splDimO(I)
        End If
    Next
'If num = 20 And perT Like "*THIS2*" Then MsgBox num & Chr(10) & perT & Chr(10) & Chr(10) & perO
    
    ' résolution des opérateurs
    res = ""
    Dim aetendre As String
    Dim listExtended() As Variant
    ReDim listExtended(LBound(spldimT) To (UBound(spldimT) - 1))
    Dim debent As String
    Dim dimension As String
    Dim error As Boolean
    error = False
    Dim THISn() As String
    Dim BEFOREn() As String
    Dim PATHn() As String
    Dim ATTRn() As String
    Dim fact As String
'If num >= 31 Then MsgBox "" & num & Chr(10) & perT & Chr(10) & perO
    For I = LBound(spldimT) To UBound(spldimT) - 1
        If Left(spldimT(I), 1) = "[" Then
            aTraiter = spldimT(I)
            ' alimentation des OPERATEURn
'If num = 22 Then MsgBox "aTraiter=" & aTraiter & Chr(10) & Chr(10) & Join(splDimO, Chr(10))
            splDim = Split(splDimO(I), ".")
            ' tout dépend de ce que ATTRn représente les attributs de la quantité ou bien l'attribut prcisé dans la définition
            '''''splAtt = Split(splDimI(i), ".")
            splAtt = Split(splDimO(I), ".")
'If UBound(splDim) <> UBound(splAtt) Then MsgBox spldimO(i) & Chr(10) & splDimI(i)
            ReDim THISn(0)
            ReDim BEFOREn(0)
            ReDim PATHn(0)
            ReDim ATTRn(0)
            For f = LBound(splDim) To UBound(splDim)
                fact = splDim(f)
'If num = 20 Then MsgBox perI & Chr(10) & "fact=" & fact
                If f = LBound(splDim) Then fact = Mid(fact, 2)
                If UBound(THISn) > 0 Then
                    ReDim Preserve THISn(1 To (f + 1))
                Else
                    ReDim THISn(1 To (f + 1))
                End If
                If UBound(BEFOREn) > 0 Then
                    ReDim Preserve BEFOREn(1 To (f + 1))
                Else
                    ReDim BEFOREn(1 To (f + 1))
                End If
                If UBound(PATHn) > 0 Then
                    ReDim Preserve PATHn(1 To (f + 1))
                Else
                    ReDim PATHn(1 To (f + 1))
                End If
                If UBound(ATTRn) > 0 Then
                    ReDim Preserve ATTRn(1 To (f + 1))
                Else
                    ReDim ATTRn(1 To (f + 1))
                End If
                If fact Like "*>*" Then
                    rfact = StrReverse(fact)
                    THISn(UBound(THISn)) = Split(StrReverse(Split(rfact, ">")(0)), ":")(0)
                    BEFOREn(UBound(BEFOREn)) = StrReverse(Split(rfact, ">")(1))
                Else
                    THISn(UBound(THISn)) = Split(fact, ":")(0)
                    BEFOREn(UBound(BEFOREn)) = ""
                End If
                PATHn(UBound(PATHn)) = fact
'If f > UBound(splAtt) Then MsgBox num & Chr(10) & perT & Chr(10) & perO & Chr(10) & UBound(splAtt) & ":" & f & "===" & splDimI(i)
                If UBound(Split(splAtt(f), ":")) > 0 Then
                    ATTRn(UBound(ATTRn)) = Split(splAtt(f), ":")(1)
                Else
                    ATTRn(UBound(ATTRn)) = ""
                End If
'If num = 20 Then MsgBox perI & Chr(10) & "fact=" & fact & Chr(10) & splAtt(f) & Chr(10) & ">>>" & ATTRn(UBound(ATTRn))
            Next
'If InStr(perT, "ATTR") And num = 12 Then MsgBox num & Chr(10) & perO & Chr(10) & perI & Chr(10) & perT & Chr(10) & Join(ATTRn, "|")
'If num = 12 And i = 0 And perT Like "*THIS2*" Then MsgBox perT & Chr(10) & perI & Chr(10) & Chr(10) & Join(THISn, Chr(10))
'If i = LBound(spldimT) And num = 20 And perT Like "*THIS2*" Then MsgBox fact & Chr(10) & Join(THISn, Chr(10))
            Dim plus As String
            Dim numn As Integer
'If num = 12 Then
    'MsgBox aTraiter & Chr(10) & perT & Chr(10) & perI & Chr(10) & perO & Chr(10) & Chr(10) & Join(THISn, Chr(10))
'End If
            For n = 1 To UBound(splDim) + 2
                If n = (UBound(splDim) + 2) Then
                    plus = ""
                    numn = 1
                Else
                    plus = "" & n
                    numn = n
                End If
'If num = 12 And i = 0 And n = 2 And perT Like "*THIS2*" Then MsgBox aTraiter
                If aTraiter Like "*THIS" & plus & "*" Then
                    aTraiter = Replace(aTraiter, "THIS" & plus, THISn(numn))
                    ' Pour enlever les attributs s'il y a THISn>suite
'If num = 12 And i = 0 And n = 2 And perT Like "*THIS2*" Then MsgBox aTraiter
                End If
                If aTraiter Like "*BEFORE" & plus & "*" Then
                    aTraiter = Replace(aTraiter, "BEFORE" & plus, BEFOREn(numn))
                    ' Pour enlever les attributs s'il y a BEFOREn>suite
                End If
                If aTraiter Like "*PATH" & plus & "*" Then
                    aTraiter = getNormalPath(Replace(aTraiter, "PATH" & plus, PATHn(numn)))
                    ' Pour enlever les attributs s'il y a PATHn>suite
                End If
                If aTraiter Like "*ATTR" & plus & "*" Then
'If num = 24 Then MsgBox plus & ":" & aTraiter & Chr(10) & ">" & ATTRn(numn) & "<"
'If numn > UBound(ATTRn) Then MsgBox numn & Chr(10) & UBound(ATTRn) & Chr(10) & aTraiter & Chr(10) & LBound(splDim) & ":" & UBound(splDim)
                    If ATTRn(numn) <> "" Then
                        aTraiter = Replace(aTraiter, "ATTR" & plus, ATTRn(numn))
                    Else
                        aTraiter = Replace(aTraiter, ":ATTR" & plus, ATTRn(numn))
                        aTraiter = Replace(aTraiter, ":OR(ATTR" & plus & ")", ATTRn(numn))
                        aTraiter = Replace(aTraiter, ":AND(ATTR" & plus & ")", ATTRn(numn))
                        aTraiter = Replace(aTraiter, ":NOT(ATTR" & plus & ")", ATTRn(numn))
                    End If
'If num = 24 Then MsgBox plus & ":" & aTraiter & Chr(10) & ">" & ATTRn(numn) & "<"
                    'If g Then
                        
                    'Else
                        'aTraiter = Replace(aTraiter, "ATTR" & plus, "")
                    'End If
'If num = 109 Then MsgBox num & Chr(10) & Join(ATTRn, "|") & Chr(10) & aTraiter
                End If
            Next
            If I = LBound(spldimT) Then jct = ""
            If I > LBound(spldimT) Then jct = "]"
            aetendre = Mid(aTraiter, 2)
'If num = 12 And i = 0 And perT Like "*THIS2*" Then MsgBox perT & Chr(10) & perI & Chr(10) & aetendre & "===" & aTraiter
'If num = 16 Then MsgBox aetendre & "===" & aTraiter
            debent = Left(aetendre, 2)
            dimension = "NOMENCLATURE"
            If Left(debent, 1) = "a" And (Right(debent, 1) = ">" Or Len(debent) = 1) Then dimension = "AREA"
            If Left(debent, 1) = "s" And (Right(debent, 1) = ">" Or Len(debent) = 1) Then dimension = "SCENARIO"
            If dimension = "NOMENCLATURE" Then
'If num = 22 Then MsgBox aetendre & Chr(10) & remplaceDoubleSup(aetendre, listN, "", 0)
'If Left(aetendre, 1) = "[" Then aetendre = Mid(aetendre, 2)
                aetendre = remplaceDoubleSup(aetendre, listN, "", 0)
'If num = 20 Then
                eten = getExtended("equ", True, aetendre, listN, num, "EQUATION", FL0.NAME, "L")
'''If num = 20 Then MsgBox aetendre & Chr(10) & UBound(eten)
'''If num = 20 Then
    '''If UBound(eten) > 0 Then MsgBox aetendre & Chr(10) & UBound(eten) & Chr(10) & Join(eten, Chr(10))
'''End If
            End If
            If dimension = "AREA" Then eten = getExtended("equ", True, aetendre, listA, num, "EQUATION", FL0.NAME, "L")
            If dimension = "SCENARIO" Then eten = getExtended("equ", True, aetendre, listS, num, "EQUATION", FL0.NAME, "L")
'If num = 109 Then MsgBox UBound(eten) & ":" & dimension & ":" & etendre & UBound(eten) & "=" & aetendre
            If UBound(eten) = 0 Then
                error = True
            Else
                listExtended(I) = eten
            End If
'If num = 12 And i = 0 And perT Like "*THIS2*" Then MsgBox perT & Chr(10) & UBound(eten) & Chr(10) & aetendre & "===" & aTraiter
'If num = 65 Then MsgBox error & ":::" & aetendre & "=ub=" & UBound(eten) & "=der=" & eten(UBound(eten))
        End If
    Next
    'If num = 22 Then MsgBox UBound(listExtended)
'If num = 12 And n = 1 Then
    'MsgBox aTraiter & Chr(10) & perT & Chr(10) & perI & Chr(10) & perO & Chr(10) & Chr(10) & Join(THISn, Chr(10))
'End If
'If num = 12 And perT Like "*THIS2*" Then MsgBox error & Chr(10) & perT & Chr(10) & perI & Chr(10) & aetendre & "===" & aTraiter
    If Not error Then
'If num = 65 Then MsgBox "not error"
'If num = 12 Then MsgBox Join(eten, Chr(10))
        Dim list0() As String
        Dim list1() As String
        Dim list2() As String
        Dim res0 As String
        Dim res1 As String
        Dim res2 As String
        Dim resf As String
        resf = ""
        list0 = listExtended(0)
'If num = 12 Then
    'MsgBox Join(list0, Chr(10))
'End If
        For i0 = LBound(list0) To UBound(list0)
            res0 = "[" & list0(i0) & "]"
            If resf = "" Then
                jct = ""
            Else
                jct = ";"
            End If
            If UBound(listExtended) > 0 Then
                list1 = listExtended(1)
'''If num = 12 Then MsgBox Join(list1, Chr(10))
                For i1 = LBound(list1) To UBound(list1)
                    res1 = "[" & list1(i1) & "]"
                    If resf = "" Then
                        jct = ""
                    Else
                        jct = ";"
                    End If
                    If UBound(listExtended) > 1 Then
                        list2 = listExtended(2)
                        For i2 = LBound(list2) To UBound(list2)
                            res2 = "[" & list2(i2) & "]"
                            If resf = "" Then
                                jct = ""
                            Else
                                jct = ";"
                            End If
                            resf = resf & jct & res0 & res1 & res2 & spldimT(UBound(spldimT))
                        Next
                    Else
                        resf = resf & jct & res2 & spldimT(UBound(spldimT))
                    End If
                Next
            Else
                resf = resf & jct & res2 & spldimT(UBound(spldimT))
            End If
        Next
        oresf = resf
        resf = isaTerme(resf, listN2, listA2, listS2, listQ2)
'If num = 224 Then MsgBox ">" & oresf & "<" & Chr(10) & ">" & resf & "<"
'If num = 12 And perT Like "*THIS2*" Then MsgBox error & Chr(10) & perT & Chr(10) & perI & Chr(10) & "oresf=" & oresf & Chr(10) & "resf=" & resf & Chr(10) & Join(listQ2, Chr(10))
        If resf = "" Then
            Dim newLogLine() As Variant
newLogLine = Array("", "ERREUR", "", "Equation", time(), "EQUATION", perT, num, "terme de l'équation inconnu1")
logNom = alimLog(logNom, newLogLine)
            resf = "ERROR:" & oresf
        End If
    Else
'If num = 65 Then MsgBox "ELSE ERROR"
        resf = "ERROR"
    End If
'If num = 11 Then MsgBox "getPerimeter end=" & resf
    getPerimeter = resf
End Function
Function getNormalPath(str As String) As String
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = True
    reg.Pattern = "(\[.*?\].*?\))"
    Dim sti As String
    Dim res As String
    Dim sep As String
    res = ""
    spl = Split(str, ".")
    For I = LBound(spl) To UBound(spl)
        sti = spl(I)
        spl2p = Split(sti, ":")
        If UBound(spl2p) > 0 Then
            spl2ps = Split(spl2p(UBound(spl2p)), ">")
            If UBound(spl2ps) > 0 Then
                sti = spl2p(LBound(spl2p)) & ">" & spl2ps(UBound(spl2ps))
            End If
        End If
        If I = LBound(spl) Then sep = ""
        If I > LBound(spl) Then sep = "."
        res = res & sep & sti
    Next
    getNormalPath = res
End Function
Function isaTerme(atester As String, listR() As String, listA() As String, listS() As String, listQ() As String) As String
' test si un terme fait partie de la liste
'If Left(atester, 16) = "[B>b1:b][a][s>s1" Then MsgBox atester & Chr(10) & Join(listR, Chr(10))
    Dim perimEt As String
    Dim atesterSansTime As String
    splQua = Split(atester, ";")
    Dim res As String
    res = ""
    For I = LBound(splQua) To UBound(splQua)
        atesteri = splQua(I)
        splAtester = Split(StrReverse(atesteri), "(")
        If UBound(splAtester) > 0 Then
            atesterSansTime = StrReverse(splAtester(1))
'If Left(atester, 16) = "[B>b1:b][a][s>s1" Then MsgBox atester & Chr(10) & atesterSansTime
            For Nolig = LBound(listR) To UBound(listR)
                perimEt = "[" & listR(Nolig) & "][" & listA(Nolig) & "][" & listS(Nolig) & "]." & listQ(Nolig)
'If Nolig < 3 Then MsgBox perimEt & Chr(10) & atesterSansTime
                If atesterSansTime = perimEt Then
                    If res = "" Then jct = ""
                    If res <> "" Then jct = ";"
                    res = res & jct & atesteri
'MsgBox perimEt & Chr(10) & atesterSansTime & Chr(10) & Res
                    Exit For
                End If
            Next
        End If
    Next
    isaTerme = res
End Function
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
Function analyseAdditivité(perO As String) As String()
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
    analyseAdditivité = res
End Function
Function setATTRIBUTS(nom As String) As Integer
    Set FLATT = g_WB_Modele.Worksheets(nom)
    With FLATT.Range("h2:h100")
        param = .VALUE
    End With
    ReDim ATTRIBUTS(0)
    For I = LBound(param, 1) To UBound(param, 1)
        If param(I, 1) = "" Then Exit For
        If UBound(ATTRIBUTS) = 0 Then
            ReDim ATTRIBUTS(1 To 1)
        Else
            ReDim Preserve ATTRIBUTS(1 To UBound(ATTRIBUTS) + 1)
        End If
        ATTRIBUTS(UBound(ATTRIBUTS)) = param(I, 1)
    Next
    setATTRIBUTS = UBound(ATTRIBUTS)
End Function
Sub setPathQua()
    'MsgBox ActiveSheet.Cells(3, 8).VALUE & Chr(10) & ActiveSheet.Cells(4, 8).VALUE
    quantityDependances.Show
End Sub
Function analyseSyntaxique(liste() As String) As String()
' Analyse syntaxique des éléments de liste
' renvoie une liste correspondante avec :
' rien si rien au départ,
' q si quantité syntaxiquement correcte,
' une expression (ex: q++q) si invalide
    ' Résultats de l'analyse
    Dim res() As String
    ReDim res(LBound(liste) To UBound(liste))
    ' join des fonctions autorisées
    Dim functionsListJoin As String
    functionsListJoin = "(" & Join(functionsList, "|") & ")"
    ' Séparateur de décimal
    Dim dsep As String
    If Application.International(xlDecimalSeparator) = "." Then
        dsep = "\."
    Else
        dsep = Application.International(xlDecimalSeparator)
    End If
    Dim reg As VBScript_RegExp_55.RegExp: Set reg = New VBScript_RegExp_55.RegExp: reg.Global = True
    Dim qua As String
    Dim qua0 As String
    For Nolig = LBound(liste) To UBound(liste)
        qua0 = liste(Nolig)
        qua = LCase(liste(Nolig))
reg.Pattern = "[ ]+": If reg.test(qua) Then qua = reg.Replace(qua, "")                      ' Les espaces
If Left(qua, 5) = "error" Then qua = ""
' Les erreurs précédentes
reg.Pattern = "\(-": If reg.test(qua) Then qua = reg.Replace(qua, "(")                      ' Quantité négative
reg.Pattern = "\[[^\]]*\]": If reg.test(qua) Then qua = reg.Replace(qua, "_")               ' Périmètre entre []
reg.Pattern = "___\.[a-zA-Z0-9_%]+": If reg.test(qua) Then qua = reg.Replace(qua, "q")      ' Les quantités
reg.Pattern = functionsListJoin & "\(": If reg.test(qua) Then qua = reg.Replace(qua, "q(")  ' Les fonctions(
reg.Pattern = functionsListJoin: If reg.test(qua) Then qua = reg.Replace(qua, "q")          ' Les fonctions
'If Nolig = UBound(liste) Then MsgBox qua
reg.Pattern = "[0-9]+[0-9\.," & dsep & "]*": If reg.test(qua) Then qua = reg.Replace(qua, "q")       ' Un nombre
'If Nolig = UBound(liste) Then MsgBox qua & Chr(10) & reg.Pattern
        ''' traitement du temps
reg.Pattern = "\(t\)": If reg.test(qua) Then qua = reg.Replace(qua, "")                     ' Le temps seul
'If Nolig = UBound(liste) Then MsgBox qua & Chr(10) & reg.Pattern
reg.Pattern = "\(t[-\+\*\/\^=][0-9\.," & dsep & "]+\)": If reg.test(qua) Then qua = reg.Replace(qua, "(t)") ' Le temps opération valeur ou bien valeur opération temps
'If Nolig = UBound(liste) Then MsgBox qua & Chr(10) & reg.Pattern
reg.Pattern = "\([0-9\.," & dsep & "]+[-\+\*\/\^]" & "t\)": If reg.test(qua) Then qua = reg.Replace(qua, "(t)")
'If Nolig = UBound(liste) Then MsgBox qua
reg.Pattern = "\$[a-z0-9_]+[\.>](last|first|ante|next)": If reg.test(qua) Then qua = reg.Replace(qua, "t") ' Le temps statique
'If Nolig = UBound(liste) Then MsgBox qua & Chr(10) & reg.Pattern
For I = 1 To 10
    reg.Pattern = "[t|q][-\+\*\/\^:=;][t|q]": If reg.test(qua) Then qua = reg.Replace(qua, "q")      ' Une plage de temps ou un calul sur les temps
    reg.Pattern = "\([t|q][-\+\*\/\^:=;][t|q]\)": If reg.test(qua) Then qua = reg.Replace(qua, "(q)") ' Une plage de temps ou un calul sur les temps
    reg.Pattern = "\([t|q]\)[-\+\*\/\^:=;]\([t|q]\)": If reg.test(qua) Then qua = reg.Replace(qua, "(q)") ' Une plage de temps ou un calul sur les temps
    reg.Pattern = "q\([t|q]\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")                  ' Le temps tout seul
Next
        ' Une opération suivie par un temps
        reg.Pattern = "[-\+\*\/=]\(t\)": If reg.test(qua) Then qua = reg.Replace(qua, "")
        ' Une opération suivie par un temps
        reg.Pattern = "[-\+\*\/=]t": If reg.test(qua) Then qua = reg.Replace(qua, "")
        ' Le temps tout seul
        reg.Pattern = "q\(t\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Le temps tout seul
        reg.Pattern = "q\(t": If reg.test(qua) Then qua = reg.Replace(qua, "q(q")
        ' Le temps tout seul
        reg.Pattern = "q\(q\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        
        ' Valeur opération quantité ou inverse
        reg.Pattern = "[0-9" & dsep & "]+[-\+\*/\^=]q": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        reg.Pattern = "q[-\+\*/\^=][0-9" & dsep & "]+": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Opération entre quantités
        reg.Pattern = "q[-\+\*/\^=]q": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Quantité négative
        reg.Pattern = "\(-q": If reg.test(qua) Then qua = reg.Replace(qua, "(q")
        ' Parenthèses vides d'une quantité
        reg.Pattern = "q\(\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Parenthèses vides seules (q ou vide?)
        reg.Pattern = "\(\)": If reg.test(qua) Then qua = reg.Replace(qua, "")
        ' Les fonctions Excel et plage de quantités
        reg.Pattern = "q\(q(;q)*\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Les fonctions Excel et plage de quantités
        'reg.Pattern = "q(;q)*": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        'reg.Pattern = "q\(q(;q)*\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' ???
        reg.Pattern = "q^[0-9]+": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ''' Les espaces
        '''reg.Pattern = "[ ]+": If reg.test(qua) Then qua = reg.Replace(qua, "")
        ' Le temps tout seul
        reg.Pattern = "t": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        For I = 1 To 10
            reg.Pattern = "\(q\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
            reg.Pattern = "q[-\+\*/\^=]q": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        Next
        ' Un nombre
        reg.Pattern = "[0-9" & dsep & "]+": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        ' Opérations entre quantités
        reg.Pattern = "q[-\+\*/\^]q": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        For I = 1 To 5
            reg.Pattern = "\(q\)": If reg.test(qua) Then qua = reg.Replace(qua, "q")
            reg.Pattern = "q[-\+\*/\^=;]q": If reg.test(qua) Then qua = reg.Replace(qua, "q")
        Next
        res(Nolig) = qua
    Next
    analyseSyntaxique = res
End Function
Sub alimFunctionsList()
    ReDim functionsList(0 To 5)
    functionsList(0) = "interpolation"
    functionsList(1) = "escalier"
    functionsList(2) = "sommeproduit"
    functionsList(3) = "somme"
    functionsList(4) = "si"
    functionsList(5) = "exp"
End Sub
Sub SetQuantites()
    '''On Error GoTo errorHandler
    openFileIfNot ("MODELE")
    Dim FL0QU As Worksheet
    Dim FLCQU As Worksheet
    Dim FLCTL As Worksheet
    Set FLCTL = Worksheets(g_CONTROL)
    Dim nameSheetEnCours As String
    Dim nameTime As String
    Dim nameArea As String
    Dim nameScenario As String
    Dim Nolig As Long
    Call alimFunctionsList
Application.StatusBar = "CONSTRUCTION DES QUANTITES"
    '''If ActiveSheet.NAME = "CONTROL" Then
        nameSheetEnCours = ActiveSheet.cbQuantites.VALUE
        nameTime = ActiveSheet.cbTime.VALUE
        nameArea = ActiveSheet.cbArea.VALUE
        nameScenario = ActiveSheet.cbScenario.VALUE
    '''Else
        '''nameSheetEnCours = ActiveSheet.NAME
    '''End If
    Set FL0QU = g_WB_Modele.Worksheets(nameSheetEnCours)
    Set FLCQU = g_WB_Modele.Worksheets("CRQUA")
    Set FLQUA = g_WB_Modele.Worksheets(g_QUANTITY)
    Set FLNOM = g_WB_Modele.Worksheets("NOMENCLATURE")
    Dim sngChrono As Single
    sngChrono = Timer
    FLQUA.Cells.Clear
    resInitGlobales = initGlobales()
    Dim LigDeb0No As Integer
    derlig = getDerLig(FLCQU)
    'derlig = Split(FLCQU.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLCQU.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    derligsheet = FLCQU.Range("A" & FLCQU.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCQU.Columns.Count).End(xlToLeft).Column
    FLCQU.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    With FLCQU.Range("a" & (LigDebCNo - 1) & ":" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Dim newLogLine() As Variant
newLogLine = Array("QUANTITES", "DEBUT", FL0QU.NAME, "GENERATION", time())
logNom = alimLog(logNom, newLogLine)

    derlig = FL0QU.Cells.SpecialCells(xlCellTypeLastCell).Row
    For Nolig = 1 To derlig
        If FL0QU.Cells(Nolig, 1) = "ACTION" Then
            LigDeb0No = Nolig + 1
            Exit For
        End If
    Next
    ' alimentation des attributs
    na = setATTRIBUTS("NOMENCLATURE")
    ' Lecture des paramètres
    With FL0QU.Range("a1:h1")
        param = .VALUE
    End With
    If ActiveSheet.NAME <> "CONTROL" Then
        For I = LBound(param, 2) To UBound(param, 2)
            If param(1, I) Like "TIME=*" Then nameTime = Split(param(1, I), "=")(1)
            If param(1, I) Like "AREA=*" Then nameArea = Split(param(1, I), "=")(1)
            If param(1, I) Like "SCENARIO=*" Then nameScenario = Split(param(1, I), "=")(1)
        Next
    End If
    If nameTime = "" Then nameTime = "TIME"
    If nameArea = "" Then nameArea = "AREA"
    If nameScenario = "" Then nameScenario = "SCENARIO"
    Set FLAREA = g_WB_Modele.Worksheets(nameArea)
    Set FLSCENARIO = g_WB_Modele.Worksheets(nameScenario)
    ' lecture de TIME
    Dim times() As String
    times = lectureTime(nameTime)
    'MsgBox Join(a, "=")
    ' lecture des areas
    Dim AREAS() As Variant
    Dim areaList() As String
    derlignom = FLAREA.Range("A" & FLAREA.Rows.Count).End(xlUp).Row
    AREAS = FLAREA.Range("A2:B" & derlignom).VALUE
    ReDim areaList(1 To UBound(AREAS, 1))
    For a = LBound(AREAS, 1) To UBound(AREAS, 1)
        areaList(a) = AREAS(a, 1)
    Next
    racineAREA = AREAS(1, 1)
    
    ' lecture des scénarios
    Dim SCENARIOS() As Variant
    Dim scenarioList() As String
    derlignom = FLSCENARIO.Range("A" & FLSCENARIO.Rows.Count).End(xlUp).Row
    SCENARIOS = FLSCENARIO.Range("A2:B" & derlignom).VALUE
    ReDim scenarioList(1 To UBound(SCENARIOS, 1))
    For a = LBound(SCENARIOS, 1) To UBound(SCENARIOS, 1)
        scenarioList(a) = SCENARIOS(a, 1)
    Next
    racineSCENARIO = SCENARIOS(1, 1)
    
    ' lecture de la nomenclature
    Dim NOMENC() As Variant
    Dim NOMENCLA() As String
    Dim IDNOMENC() As Variant
    Dim IDNOMENCLA() As String
    Dim eten() As String
    ReDim eten(0)
    derlignom = FLNOM.Range("C" & FLNOM.Rows.Count).End(xlUp).Row
    NOMENC = FLNOM.Range("c2:c" & derlignom).VALUE
    IDNOMENC = FLNOM.Range("f2:f" & derlignom).VALUE
'MsgBox LBound(NOMENC, 1) & ":" & UBound(NOMENC, 1)
    ReDim NOMENCLA(1 To UBound(NOMENC, 1))
    ReDim IDNOMENCLA(1 To UBound(IDNOMENC, 1))
    For nol = LBound(NOMENC, 1) To UBound(NOMENC, 1)
        NOMENCLA(nol) = NOMENC(nol, LBound(NOMENC, 2))
        IDNOMENCLA(nol) = IDNOMENC(nol, LBound(IDNOMENC, 2))
    Next
'MsgBox Join(NOMENCLA, "@")
    'Alimentation des tableaux
    Dim Cla() As Variant
    derligsheet = FL0QU.Cells.SpecialCells(xlCellTypeLastCell).Row
    DerColSheet = Cells(LigDeb0No - 1, Columns.Count).End(xlToLeft).Column
    With FL0QU.Range("a" & LigDeb0No & ":" & "AP" & derligsheet)
        ReDim Cla(1 To (derligsheet - LigDeb0No), 1 To DerColSheet)
        Cla = .VALUE
    End With
    ' boucle principale
    Dim lig As Integer
    Dim entite As String
    Dim error As Boolean
    Dim listRes() As String
    Dim listSce() As String
    Dim listAre() As String
    Dim listQua() As String
    Dim listTim() As String
    Dim listEqu() As String
    Dim listPer() As String
    Dim listNam() As String
    Dim listSCALE() As String
    Dim listUNIT() As String
    'Dim listFORMULA() As String
    Dim listSUBSTITUTE() As String
    Dim listDEFAUT() As String
    Dim listVALUE() As String
    Dim listBoucle() As String
    ReDim listRes(0)
    ReDim listSce(0)
    ReDim listAre(0)
    ReDim listQua(0)
    ReDim listTim(0)
    ReDim listEqu(0)
    ReDim listPer(0)
    ReDim listNam(0)
    ReDim listSCALE(0)
    ReDim listUNIT(0)
    'ReDim listFORMULA(0)
    ReDim listSUBSTITUTE(0)
    ReDim listDEFAUT(0)
    ReDim listVALUE(0)
    ReDim listBoucle(0)
    Dim listResNew() As String
    Dim listSceNew() As String
    Dim listAreNew() As String
    Dim listQuaNew() As String
    Dim listTimNew() As String
    Dim listEquNew() As String
    Dim listPerNew() As String
    Dim listNamNew() As String
    Dim listSCALENew() As String
    Dim listUNITNew() As String
    'Dim listFORMULANew() As String
    Dim listSUBSTITUTENew() As String
    Dim listDEFAUTNew() As String
    Dim listVALUENew() As String
    Dim listBOUCLENew() As String
    ReDim listResNew(0)
    ReDim listSceNew(0)
    ReDim listAreNew(0)
    ReDim listQuaNew(0)
    ReDim listTimNew(0)
    ReDim listEquNew(0)
    ReDim listPerNew(0)
    ReDim listNamNew(0)
    ReDim listSCALENew(0)
    ReDim listUNITNew(0)
    'ReDim listFORMULANew(0)
    ReDim listSUBSTITUTENew(0)
    ReDim listDEFAUTNew(0)
    ReDim listVALUENew(0)
    ReDim listBOUCLENew(0)
    Dim ind As Integer
    ind = 0
    Dim numlig() As String
    ReDim numlig(0)
    Dim numLigNew() As String
    ReDim numLigNew(0)
    Dim areaR As String
    Dim scenarioR As String
    Dim nextLigne As Integer
    Dim nbAjout As Integer
    Dim listSurcharge() As Boolean
    Dim ligSurcharge As String
    ligSurcharge = ""
    Dim cibSurcharge As String
    cibSurcharge = ""
    Dim ligRetrait As String
    ligRetrait = ""
    Dim cibRetrait As String
    cibRetrait = ""
    Dim nbSurcharge As Integer
    nbSurcharge = 0
    Dim listMinus() As Boolean
    Dim nbRealMinus As Integer
    Dim qualue As String
    Dim timelu As String
    Dim equalu As String
    Dim perimetre As String
    Dim nomlu As String
    Dim SCALElu As String
    Dim UNITlu As String
    'Dim FORMULAlu As String
    Dim SUBSTITUTElu As String
    Dim DEFAUTlu As String
    Dim VALUElu As String
    Dim BOUCLElu As String
    Dim nbLigOK As Integer
    Dim listNb() As String
    ReDim listNb(1 To UBound(Cla, 1), 1 To 1)
    Dim nbRetraitsEnTout As Integer
    nbRetraitsEnTout = 0
    Dim compteur As Integer
    Dim lastCompteur As Integer
Application.StatusBar = "CONSTRUCTION DES QUANTITES : GENERATION : 0 / " & UBound(Cla, 1) & " " & compteur & " %"
    Dim beginNum As Integer
    beginNum = 0
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        If Trim(Cla(Nolig, 1)) = "BEGIN" Then
            beginNum = Nolig
            Exit For
        End If
    Next
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        listNb(Nolig, 1) = ""
'If Nolig = 49 Then MsgBox Nolig
        If Trim(Cla(Nolig, 1)) = "END" Then Exit For
        If Nolig > beginNum And Left(Cla(Nolig, 1), 1) <> "#" And Trim(Cla(Nolig, 1) & Cla(Nolig, 2) & Cla(Nolig, 3) & Cla(Nolig, 4) & Cla(Nolig, 5)) <> "" Then
compteur = 100 * Nolig / UBound(Cla, 1)
'If Int(Round(0.5 + compteur)) <> lastCompteur Then
Application.StatusBar = "CONSTRUCTION DES QUANTITES : GENERATION : " & Nolig & " / " & UBound(Cla, 1) & " " & compteur & " %"
'FLCTL.Cells(7, 7).VALUE = compteur & " %"
'lastCompteur = compteur
'End If
            error = False
            lig = Nolig
'If Nolig = 164 Then(LigDeb0No + Nolig - 1) = 164(LigDeb0No + Nolig - 1) > 110 And (LigDeb0No + Nolig - 1) < 120 And
    'MsgBox entite
'End If
            entite = jonction(lig, Cla)
'If (LigDeb0No + Nolig - 1) = 164 Then
    'MsgBox entite & Chr(10) & Cla(Nolig, 2) & Chr(10) & Cla(Nolig, 3) & Chr(10) & Cla(Nolig, 4) & Chr(10) & Cla(Nolig, 5) & Chr(10) & Cla(Nolig, 6) & Chr(10) & Cla(Nolig, 7) & Chr(10) & Cla(Nolig, 8)
'End If
            qualue = Trim(Cla(Nolig, quantity))
            timelu = Trim(Cla(Nolig, temps))
            equalu = Trim(Cla(Nolig, equation))
            nomlu = Trim(Cla(Nolig, NAME))
            'SCALE   UNIT    TIME    FORMULA SUBSTITUTE  DEFAUT  VALUE   BOUCLE
            SCALElu = Trim(Cla(Nolig, SCALE0))
            UNITlu = Trim(Cla(Nolig, UNIT))
            'FORMULAlu = Trim(Cla(Nolig, Formula))
            SUBSTITUTElu = Trim(Cla(Nolig, SUBSTITUTE))
            DEFAUTlu = Trim(Cla(Nolig, DEFAUT))
            VALUElu = Trim(Cla(Nolig, VALUE))
            BOUCLElu = Trim(Cla(Nolig, BOUCLE))
            'If timelu = "" Then MsgBox Nolig & ""
            ' résolution des raccourcis
            If qualue = "" Or TestNotAlphaQua(qualue) Then
            where = "QUANTITY"
newLogLine = Array("", "ERREUR", FL0QU.NAME, "Quantité", time(), where, qualue, LigDeb0No + Nolig - 1, "quantité mal formée")
logNom = alimLog(logNom, newLogLine)
                GoTo continueLoop
            End If
            entite = remplaceDoubleSup(entite, NOMENCLA, FL0QU.NAME, (LigDeb0No + Nolig - 1))
'If Nolig = 275 Then MsgBox entite
            If entite = "" And Trim(Cla(Nolig, ID)) = "" Then GoTo continueLoop
            nbLigOK = nbLigOK + 1
            If entite Like "*>>*" Then
                splEntite = Split(entite, ".")
                where = ""
                For I = LBound(splEntite) To UBound(splEntite)
                    If splEntite(I) Like "*>>*" Then where = where & "," & "BEFORE" & (I + 1)
                Next
                where = Mid(where, 2)
newLogLine = Array("", "ERREUR", FL0QU.NAME, "Raccourci", time(), where, entite, LigDeb0No + Nolig - 1, "raccourci non résolu")
logNom = alimLog(logNom, newLogLine)
                error = True
            End If
            If error Then
                listNb(Nolig, 1) = "0"
                GoTo continueLoop
            End If
'''MsgBox (LigDeb0No + Nolig - 1) & ">>>" & Trim(Cla(Nolig, 2)) & "<<<"
            If Trim(Cla(Nolig, ID)) <> "" Then
                ' Traitement des ID
                nid = 0
                For I = LBound(IDNOMENCLA) To UBound(IDNOMENCLA)
                    If Trim(IDNOMENCLA(I)) = Trim(Cla(Nolig, 2)) Then
                        nid = nid + 1
                        If UBound(eten) = 0 Then ReDim eten(1 To nid)
                        If UBound(eten) > 0 Then ReDim Preserve eten(1 To nid)
                        eten(UBound(eten)) = Trim(NOMENCLA(I))
                    End If
                Next
                If nid = 0 Then
newLogLine = Array("", "ERREUR", FL0QU.NAME, "Extension", time(), "ID", Trim(Cla(Nolig, 2)), (LigDeb0No + Nolig - 1), "ID inconnu")
logNom = alimLog(logNom, newLogLine)
                    listNb(Nolig, 1) = "0"
                    GoTo continueLoop
                End If
            Else
                ' résolution des facteurs
'If (LigDeb0No + Nolig - 1) = 8 Then MsgBox "NOMENCLA=" & Join(NOMENCLA, "||")
'MsgBox "1 " & entite
'If Nolig = 49 Then MsgBox (LigDeb0No + Nolig - 1)
                eten = getExtended("per", True, entite, NOMENCLA, LigDeb0No + Nolig - 1, "QUANTITE", FL0QU.NAME, "L")
'If Nolig = 49 Then MsgBox UBound(eten)
            End If
            listNb(Nolig, 1) = "" & UBound(eten)
            If UBound(eten) > 0 Then
                isat = isATime(timelu, times)
'If (LigDeb0No + Nolig - 1) = 16 Then MsgBox isat & ":::" & Join(times, "|")
                If Not isat Then 'And timelu <> "" Then
                    where = "TEMPS"
newLogLine = Array("", "ERREUR", FL0QU.NAME, "Temps", time(), where, timelu, LigDeb0No + Nolig - 1, "Temps inconnu")
logNom = alimLog(logNom, newLogLine)
                    error = True
                    listNb(Nolig, 1) = ""
                    GoTo continueLoop
                End If
                ' AREA
                areaR = Trim(Cla(Nolig, AREA))
                If areaR = "" Then areaR = racineAREA
                etena = getExtended("per", True, areaR, areaList, LigDeb0No + Nolig - 1, "AREA", FL0QU.NAME, "L")
'MsgBox (LigDeb0No + Nolig - 1) & ">>>" & UBound(etena)
                If UBound(etena) = 0 Then listNb(Nolig, 1) = "0"
                If UBound(etena) > 0 Then
                    scenarioR = Trim(Cla(Nolig, SCENARIO))
                    If scenarioR = "" Then scenarioR = racineSCENARIO
                    etens = getExtended("per", True, scenarioR, scenarioList, LigDeb0No + Nolig - 1, "SCENARIO", FL0QU.NAME, "L")
'MsgBox entite & ":::" & (LigDeb0No + Nolig - 1) & ": e=" & UBound(eten) & ": a=" & UBound(etena) & " s=" & UBound(etens)
'If (LigDeb0No + Nolig - 1) = 16 Then MsgBox (LigDeb0No + Nolig - 1) & "<<<>>>" & Join(eten, "|")
                    '''ReDim listSurcharge(1 To (UBound(eten) * UBound(etena) * UBound(etens)))
                    nextLigne = 0
                    nbAjout = 0
                    If UBound(etens) = 0 Then listNb(Nolig, 1) = "0"
                    If UBound(etens) > 0 Then
                        listNb(Nolig, 1) = "" & (UBound(eten) * UBound(etena) * UBound(etens))
                        oldUb = UBound(listRes)
                        perimetre = "[" & entite & "][" & areaR & "][" & scenarioR & "]"
                        ReDim listSurcharge(1 To (UBound(eten) * UBound(etena) * UBound(etens)))
                        ' Surcharge
                        If Trim(Cla(Nolig, ACTION)) <> "minus" Then
                            For ne = LBound(eten) To UBound(eten)
                                For na = LBound(etena) To UBound(etena)
                                    For ns = LBound(etens) To UBound(etens)
                                        nextLigne = nextLigne + 1
'If (LigDeb0No + Nolig - 1) = 97 Then MsgBox UBound(listSurcharge) & ":::" & eten(ne)
                                        listSurcharge(nextLigne) = False
                                        nbAjout = nbAjout + 1
                                        If UBound(numlig) > 0 Then
                                            listSur = ""
                                            ' on boucle jusqu'à la ligne en cours
                                            For I = LBound(numlig) To UBound(numlig)
                                                '''''If (LigDeb0No + Nolig - 1) < numLig(i) Then
'If (LigDeb0No + Nolig - 1) = 9 Then MsgBox i & ":" & numLig(i) & eten(ne) & "|" & etena(na) & "|" & etens(ns) & "|" & qualue & "|" & timelu & ":::" & listRes(i) & "|" & listAre(i) & "|" & listSce(i) & "|" & listQua(i) & "|" & listTim(i)
    If eten(ne) = listRes(I) And etena(na) = listAre(I) And etens(ns) = listSce(I) And qualue = listQua(I) And timelu = listTim(I) Then
'If (LigDeb0No + Nolig - 1) = 9 Then MsgBox i & ":" & numLig(i) & ":::" & nbSurcharge
                                                    ' surcharge ici
'If (LigDeb0No + Nolig - 1) = 9 Then MsgBox i & ":" & numLig(i) & ":::" & nbSurcharge
'If (LigDeb0No + Nolig - 1) = 97 Then MsgBox UBound(listSurcharge) & ":OK:" & eten(ne)



                                                    listSurcharge(nextLigne) = True
                                                    nbAjout = nbAjout - 1
                                                    nbSurcharge = nbSurcharge + 1
                                                    
                                                    
                                                    If ligSurcharge <> "" Then lastLigSurcharge = Split(ligSurcharge, ";")(UBound(Split(ligSurcharge, ";")))
                                                    If ligSurcharge = "" Then lastLigSurcharge = ""
                                                    If lastLigSurcharge <> ("" & (LigDeb0No + Nolig - 1)) Then
                                                        ligSurcharge = ligSurcharge & ";" & (LigDeb0No + Nolig - 1)
                                                    End If
                                                    If listSur <> "" Then listSur = listSur & "," & numlig(I)
                                                    If listSur = "" Then listSur = "" & numlig(I)
                                                    
                                                    
                                                    
                                                    
                                                    '''ligSurcharge = ligSurcharge & ";" & (LigDeb0No + Nolig - 1)
                                                    '''If listSur <> "" Then listSur = listSur & "," & numLig(i)
                                                    '''If listSur = "" Then listSur = "" & numLig(i)
                                                    numlig(I) = numlig(I) & "," & (LigDeb0No + Nolig - 1)
                                                    ' ajout surcharge des colonnes
                                                    listEqu(I) = equalu
                                                    '''listPer(oldUb + nextLigneAjout) = perimetre
                                                    listNam(I) = nomlu
                                                    listSCALE(I) = SCALElu
                                                    listUNIT(I) = UNITlu
                                                    listSUBSTITUTE(I) = SUBSTITUTElu
                                                    listDEFAUT(I) = DEFAUTlu
                                                    listVALUE(I) = VALUElu
                                                    listBoucle(I) = BOUCLElu
                                                    '''cibSurcharge = cibSurcharge & "," & numLig(i)
                                                    Exit For
    End If
                                                '''''End If
                                            Next
                                            '''If listSur <> "" Then cibSurcharge = cibSurcharge & ";" & listSur
                                            If listSur <> "" Then
                                                If lastLigSurcharge <> ("" & (LigDeb0No + Nolig - 1)) Then
                                                    'lastListSur = Split(listSur, ";")(0)
                                                    cibSurcharge = cibSurcharge & ";" & listSur
                                                End If
                                            End If
                                        End If
                                    Next
                                Next
                            Next
                        End If
                        If Trim(Cla(Nolig, ACTION)) = "minus" Then
                            ReDim listMinus(LBound(numlig) To UBound(numlig))
                            nextLigne = 0
                            nbMinus = 0
                            nbRealMinus = 0
                            listCibEnCours = ""
                            For I = LBound(numlig) To UBound(numlig)
                                listMinus(I) = False
                            Next
                            For ne = LBound(eten) To UBound(eten)
                                For na = LBound(etena) To UBound(etena)
                                    For ns = LBound(etens) To UBound(etens)
                                        nextLigne = nextLigne + 1
                                        nbMinus = nbMinus + 1
'aaa = eten(ne) & "|" & etena(na) & "|" & etens(ns) & "|" & qualue & "|" & timelu
                                        For I = LBound(numlig) To UBound(numlig)
'bbb = listRes(i) & "|" & listAre(i) & "|" & listSce(i) & "|" & listQua(i) & "|" & listTim(i)
'MsgBox aaa & "===" & bbb
    If eten(ne) = listRes(I) And etena(na) = listAre(I) And etens(ns) = listSce(I) And qualue = listQua(I) And timelu = listTim(I) Then
                                                listMinus(I) = True
                                                '''ligRetrait = ligRetrait & "," & (LigDeb0No + Nolig - 1)
                                                '''cibRetrait = cibRetrait & "," & numLig(i)
                                                
                                                Dim lastRetrait As String
                                                If ligRetrait <> "" Then
                                                    lastRetrait = Split(ligRetrait, ";")(UBound(Split(ligRetrait, ";")))
                                                    If lastRetrait <> ("" & (LigDeb0No + Nolig - 1)) Then ligRetrait = ligRetrait & ";" & (LigDeb0No + Nolig - 1)
                                                End If
                                                If ligRetrait = "" Then ligRetrait = "" & (LigDeb0No + Nolig - 1)
                                                If listCibEnCours <> "" Then
                                                    lastRetrait = Split(listCibEnCours, ";")(UBound(Split(listCibEnCours, ";")))
                                                    If lastRetrait <> ("" & numlig(I)) Then listCibEnCours = listCibEnCours & "," & numlig(I)
                                                End If
                                                If listCibEnCours = "" Then listCibEnCours = "" & numlig(I)
                                                nbRealMinus = nbRealMinus + 1
                                                nbRetraitsEnTout = nbRetraitsEnTout + 1
                                                Exit For
    End If
                                        Next
                                    Next
                                Next
                            Next
                            
                            If nbRealMinus > 0 Then
                                If cibRetrait <> "" Then cibRetrait = cibRetrait & ";" & listCibEnCours
                                If cibRetrait = "" Then cibRetrait = listCibEnCours
                                ReDim listResNew(1 To (UBound(listRes) - nbRealMinus))
                                ReDim listSceNew(1 To (UBound(listSce) - nbRealMinus))
                                ReDim listAreNew(1 To (UBound(listAre) - nbRealMinus))
                                ReDim listQuaNew(1 To (UBound(listQua) - nbRealMinus))
                                ReDim listTimNew(1 To (UBound(listTim) - nbRealMinus))
                                ReDim listEquNew(1 To (UBound(listEqu) - nbRealMinus))
                                ReDim listPerNew(1 To (UBound(listPer) - nbRealMinus))
                                ReDim listNamNew(1 To (UBound(listNam) - nbRealMinus))
                                ReDim listSCALENew(1 To (UBound(listSCALE) - nbRealMinus))
                                ReDim listUNITNew(1 To (UBound(listUNIT) - nbRealMinus))
                                'ReDim listFORMULANew(1 To (UBound(listFORMULA) - nbRealMinus))
                                ReDim listSUBSTITUTENew(1 To (UBound(listSUBSTITUTE) - nbRealMinus))
                                ReDim listDEFAUTNew(1 To (UBound(listDEFAUT) - nbRealMinus))
                                ReDim listVALUENew(1 To (UBound(listVALUE) - nbRealMinus))
                                ReDim listBOUCLENew(1 To (UBound(listBoucle) - nbRealMinus))
                                ReDim numLigNew(1 To (UBound(numlig) - nbRealMinus))
                                Dim j As Integer
                                j = 0
                                For I = LBound(listRes) To UBound(listRes)
                                    If Not listMinus(I) Then
                                        j = j + 1
                                        listResNew(j) = listRes(I)
                                        listSceNew(j) = listSce(I)
                                        listAreNew(j) = listAre(I)
                                        listQuaNew(j) = listQua(I)
                                        listTimNew(j) = listTim(I)
                                        listEquNew(j) = listEqu(I)
                                        listPerNew(j) = listPer(I)
                                        listNamNew(j) = listNam(I)
                                        listSCALENew(j) = listSCALE(I)
                                        listUNITNew(j) = listUNIT(I)
                                        'listFORMULANew(j) = listFORMULA(i)
                                        listSUBSTITUTENew(j) = listSUBSTITUTE(I)
                                        listDEFAUTNew(j) = listDEFAUT(I)
                                        listVALUENew(j) = listVALUE(I)
                                        listBOUCLENew(j) = listBoucle(I)
                                        numLigNew(j) = numlig(I)
                                    End If
                                Next
                                ReDim listRes(1 To UBound(listResNew))
                                ReDim listSce(1 To UBound(listSceNew))
                                ReDim listAre(1 To UBound(listAreNew))
                                ReDim listQua(1 To UBound(listQuaNew))
                                ReDim listTim(1 To UBound(listTimNew))
                                ReDim listEqu(1 To UBound(listEquNew))
                                ReDim listPer(1 To UBound(listPerNew))
                                ReDim listNam(1 To UBound(listNamNew))
                                ReDim listSCALE(1 To UBound(listSCALENew))
                                ReDim listUNIT(1 To UBound(listUNITNew))
                                'ReDim listFORMULA(1 To UBound(listFORMULANew))
                                ReDim listSUBSTITUTE(1 To UBound(listSUBSTITUTENew))
                                ReDim listDEFAUT(1 To UBound(listDEFAUTNew))
                                ReDim listVALUE(1 To UBound(listVALUENew))
                                ReDim listBoucle(1 To UBound(listBOUCLENew))
                                ReDim numlig(1 To UBound(numLigNew))
                                For I = LBound(listResNew) To UBound(listResNew)
                                    listRes(I) = listResNew(I)
                                    listSce(I) = listSceNew(I)
                                    listAre(I) = listAreNew(I)
                                    listQua(I) = listQuaNew(I)
                                    listTim(I) = listTimNew(I)
                                    listEqu(I) = listEquNew(I)
                                    listPer(I) = listPerNew(I)
                                    listNam(I) = listNamNew(I)
                                    listSCALE(I) = listSCALENew(I)
                                    listUNIT(I) = listUNITNew(I)
                                    'listFORMULA(i) = listFORMULANew(i)
                                    listSUBSTITUTE(I) = listSUBSTITUTENew(I)
                                    listDEFAUT(I) = listDEFAUTNew(I)
                                    listVALUE(I) = listVALUENew(I)
                                    listBoucle(I) = listBOUCLENew(I)
                                    numlig(I) = numLigNew(I)
                                Next
                                listNb(Nolig, 1) = "-" & nbRealMinus
                            Else
newLogLine = Array("", "ALERTE", FL0QU.NAME, "", "", "", entite, (LigDeb0No + Nolig - 1), "Retrait non résolu")
logNom = alimLog(logNom, newLogLine)
                            End If
                        Else
'If (LigDeb0No + Nolig - 1) = 97 Then MsgBox ":SUITE:" & UBound(eten) & ":nbAjout=" & nbAjout
                            If ind = 0 Then
                                ReDim listRes(1 To (UBound(listRes) + nbAjout))
                                ReDim listSce(1 To (UBound(listSce) + nbAjout))
                                ReDim listAre(1 To (UBound(listAre) + nbAjout))
                                ReDim listQua(1 To (UBound(listQua) + nbAjout))
                                ReDim listTim(1 To (UBound(listTim) + nbAjout))
                                ReDim listEqu(1 To (UBound(listEqu) + nbAjout))
                                ReDim listPer(1 To (UBound(listPer) + nbAjout))
                                ReDim listNam(1 To (UBound(listNam) + nbAjout))
                                ReDim listSCALE(1 To (UBound(listSCALE) + nbAjout))
                                ReDim listUNIT(1 To (UBound(listUNIT) + nbAjout))
                                'ReDim listFORMULA(1 To (UBound(listFORMULA) + nbAjout))
                                ReDim listSUBSTITUTE(1 To (UBound(listSUBSTITUTE) + nbAjout))
                                ReDim listDEFAUT(1 To (UBound(listDEFAUT) + nbAjout))
                                ReDim listVALUE(1 To (UBound(listVALUE) + nbAjout))
                                ReDim listBoucle(1 To (UBound(listBoucle) + nbAjout))
                                ReDim numlig(1 To (UBound(numlig) + nbAjout))
                            Else
                                ReDim Preserve listRes(1 To (UBound(listRes) + nbAjout))
                                ReDim Preserve listSce(1 To (UBound(listSce) + nbAjout))
                                ReDim Preserve listAre(1 To (UBound(listAre) + nbAjout))
                                ReDim Preserve listQua(1 To (UBound(listQua) + nbAjout))
                                ReDim Preserve listTim(1 To (UBound(listTim) + nbAjout))
                                ReDim Preserve listEqu(1 To (UBound(listEqu) + nbAjout))
                                ReDim Preserve listPer(1 To (UBound(listPer) + nbAjout))
                                ReDim Preserve listNam(1 To (UBound(listNam) + nbAjout))
                                ReDim Preserve listSCALE(1 To (UBound(listSCALE) + nbAjout))
                                ReDim Preserve listUNIT(1 To (UBound(listUNIT) + nbAjout))
                                'ReDim Preserve listFORMULA(1 To (UBound(listFORMULA) + nbAjout))
                                ReDim Preserve listSUBSTITUTE(1 To (UBound(listSUBSTITUTE) + nbAjout))
                                ReDim Preserve listDEFAUT(1 To (UBound(listDEFAUT) + nbAjout))
                                ReDim Preserve listVALUE(1 To (UBound(listVALUE) + nbAjout))
                                ReDim Preserve listBoucle(1 To (UBound(listBoucle) + nbAjout))
                                ReDim Preserve numlig(1 To (UBound(numlig) + nbAjout))
                            End If
                            ind = ind + nbAjout
                            nextLigne = 0
                            nextLigneAjout = 1
                            For ne = LBound(eten) To UBound(eten)
                                For na = LBound(etena) To UBound(etena)
                                    For ns = LBound(etens) To UBound(etens)
                                        nextLigne = nextLigne + 1
'If (LigDeb0No + Nolig - 1) = 97 Then MsgBox eten(ne) & "<>" & LBound(listSurcharge) & ":::" & nextLigne & ":" & listSurcharge(nextLigne)
                                        If Not listSurcharge(nextLigne) Then
'If (LigDeb0No + Nolig - 1) = 99 Then MsgBox nextLigne
'If (oldUb + nextLigne) > UBound(listRes) Then
    'MsgBox scenarioR & ":" & entite & (LigDeb0No + Nolig - 1) & "::e=" & UBound(eten) & ":a=" & UBound(etena) & ":s=" & UBound(etens) & ":::aj=" & nbAjout
    'MsgBox UBound(listRes) & "<" & (oldUb + nextLigne) & "    nextline=" & nextLigne & "::: ne=" & ne
'End If
                                            listRes(oldUb + nextLigneAjout) = eten(ne)
                                            listAre(oldUb + nextLigneAjout) = etena(na)
                                            listSce(oldUb + nextLigneAjout) = etens(ns)
                                            listQua(oldUb + nextLigneAjout) = qualue
                                            listTim(oldUb + nextLigneAjout) = timelu
                                            listEqu(oldUb + nextLigneAjout) = equalu
                                            listPer(oldUb + nextLigneAjout) = perimetre
                                            listNam(oldUb + nextLigneAjout) = nomlu
                                            listSCALE(oldUb + nextLigneAjout) = SCALElu
                                            listUNIT(oldUb + nextLigneAjout) = UNITlu
                                            'listFORMULA(oldUb + nextLigneAjout) = FORMULAlu
                                            listSUBSTITUTE(oldUb + nextLigneAjout) = SUBSTITUTElu
                                            listDEFAUT(oldUb + nextLigneAjout) = DEFAUTlu
                                            listVALUE(oldUb + nextLigneAjout) = VALUElu
                                            listBoucle(oldUb + nextLigneAjout) = BOUCLElu
                                            numlig(oldUb + nextLigneAjout) = "" & (LigDeb0No + Nolig - 1)
                                            nextLigneAjout = nextLigneAjout + 1
                                        End If
                                        
                                    Next
                                Next
                            Next
                        End If
                    End If
                End If
            End If
        End If
continueLoop:
    Next
    
    ' Deuxième boucle pour résoudre les équations
    Dim listEquations() As String
    Dim perimEt As String
    Dim termeEtendu As String
    Dim equat As String
    Dim equationAtraiter As String
    Dim equationAtraiterEnCours As String
    Dim equationAtraiterInit As String
    Dim jcts As String
    Dim suite As String
    Dim apt0 As String
    compteur = 0
    lastCompteur = 0
    Dim suiteOK As Boolean
    Dim doubleatraiter As String
    Dim ainjecter As String
    Dim sheet2injecter As Integer
    Dim chaine As String
    Dim chaineWithout As String
    Dim resu() As String
    Dim resuWithout() As String
    Dim resuget() As String
    Dim resugetboucle As String
    Dim nu As Integer
    Dim valeurboucle As String
    Dim aj As String
    Dim instance As String
    Dim instanceList() As String
    Dim chaineSansAtt As String
    'Dim chaineDeBouclage As String
    For Nolig = 1 To UBound(listRes)
        equat = listEqu(Nolig)
'MsgBox nolog & ":" & equat
        compteur = Int(100 * Nolig / UBound(listRes))
        If compteur <> lastCompteur Then
Application.StatusBar = "CONSTRUCTION DES QUANTITES : EQUATION : " & Nolig & " / " & UBound(listRes) & " " & compteur & " %"
lastCompteur = compteur
        End If
        suiteOK = True
        If Trim(listBoucle(Nolig)) <> "" Then
            '' Traitement des BOUCLAGES
            ' modif 20/09/16
            ' calcul de la liste de bouclage s'appuie sur le périmètre générique
            splQ = Split(listRes(Nolig), ".")
            ' calcul du numéro du facteur pour le LAST ou FIRST, 0 sinon
            lastc = "0"
            If InStr(Trim(listBoucle(Nolig)), "LAST") > 0 Then
                lastc = Right(Trim(listBoucle(Nolig)), 1)
                If lastc = "T" Then lastc = "" & (UBound(splQ) + 1)
            End If
            If InStr(Trim(listBoucle(Nolig)), "FIRST") > 0 Then
                lastc = Right(Trim(listBoucle(Nolig)), 1)
                If lastc = "T" Then lastc = "" & (UBound(splQ) + 1)
            End If
            ReDim resu(LBound(splQ) To UBound(splQ))
            ReDim resuWithout(LBound(splQ) To UBound(splQ))
            If lastc <> "0" Then
                ' cas FIRST ou LAST
                ' ??? cas ou FIRST ou LAST sans n et avec un EACH=> prendre le facteur du EACH ???
            Else
                ' cas autre que FIRST ou LAST
                If InStr(listBoucle(Nolig), ">") > 0 Then
                    ' cas ou le bouclage porte sur un chemin
'If numlig(Nolig) = 59 Then MsgBox Nolig & ":" & equat & ":" & Trim(listBOUCLE(Nolig))
                    For t = LBound(splQ) To UBound(splQ)
'If numlig(Nolig) = 59 Then MsgBox Nolig & ":" & equat & ":" & Trim(listBOUCLE(Nolig)) & Chr(10) & Right(splQ(t), Len(Trim(listBOUCLE(Nolig))))
                        'chaineSansAtt = Split(Right(splQ(t), Len(Trim(listBOUCLE(Nolig)))) & ":", ":")(0)
                        chaineSansAtt = Split(splQ(t) & ":", ":")(0)
                        chaineSansAtt = Right(chaineSansAtt, Len(Trim(listBoucle(Nolig))))
'If numlig(Nolig) = 59 Then MsgBox Nolig & ":" & equat & ":" & Trim(listBOUCLE(Nolig)) & Chr(10) & chaineSansAtt
                        If chaineSansAtt = Trim(listBoucle(Nolig)) Then
                            lastc = "" & (t + 1)
                            Exit For
                        End If
                    Next
'If numlig(Nolig) = 59 Then MsgBox Nolig & ":" & equat & ":" & Trim(listBOUCLE(Nolig)) & Chr(10) & Right(splQ(t), Len(Trim(listBOUCLE(Nolig))))
                Else
                    ' cas ou le bouclage porte sur un concept simple ==> dernier facteur
                    acomparer = Split(Right(splQ(UBound(splQ)), InStr(StrReverse(splQ(UBound(splQ))), ">") - 1), ":")(0)
                    If acomparer = Trim(listBoucle(Nolig)) Then lastc = "" & (UBound(splQ) + 1)
                    '''If Right(splQ(UBound(splQ)), Len(Trim(listBOUCLE(Nolig)))) = Trim(listBOUCLE(Nolig)) Then lastc = "" & (UBound(splQ) + 1)
                End If
            End If
'If numlig(Nolig) = 59 Then MsgBox Nolig & ":" & lastc & ":" & Trim(listBOUCLE(Nolig))
            If lastc <> "0" Then
            ' le facteur de bouclage a été déterminé
            For t = LBound(splQ) To UBound(splQ)
                If lastc = "" & (t + 1) Then
                    chaineWithout = Mid(splQ(t), 1, Len(splQ(t)) - InStr(StrReverse(splQ(t)), ">")) & ">"
                    ' Détermination du facteur génarique de bouclage
                    chaine = Split(Split(Split(listPer(Nolig), "][")(0), "[")(1), ".")(t)
                    If InStr(chaine, "EACH") > 0 Or InStr(chaine, "LEAF") > 0 Or InStr(chaine, "DESC") > 0 Then
                        ' Cas où la définition de la quantité contient un opérateur factorisant
                    Else
                        ' Sinon par défaut c'est EACH
                        chaine = chaineWithout & "EACH"
                    End If
                Else
                    chaine = splQ(t)
                    chaineWithout = splQ(t)
                End If
                resu(t) = chaine
                resuWithout(t) = chaineWithout
            Next
            chaine = Join(resu, ".")
            chaineWithout = Join(resuWithout, ".")
            ' Liste des items possibles de la somme du bouclage
'MsgBox Nolig & Chr(10) & listRes(Nolig) & Chr(10) & listPer(Nolig) & Chr(10) & chaine
            '''chaineDeBouclage = Split(Split(listPer(Nolig) & "]", "]")(0), "[")(1)
            ' Liste des items possibles de la somme du bouclage
'MsgBox Nolig & Chr(10) & listRes(Nolig) & Chr(10) & listPer(Nolig) & Chr(10) & chaine
            '''resuget = getExtended("per", True, chaineDeBouclage, NOMENCLA, CInt(numlig(Nolig)), "QUANTITE", FLQUA.NAME, "L")
            resuget = getExtended("per", True, chaine, NOMENCLA, CInt(numlig(Nolig)), "QUANTITE", FLQUA.NAME, "L")
            'resuget = getExtended("per", True, chaine, NOMENCLA, CInt(numlig(Nolig)), "QUANTITE", FLQUA.NAME, "L")
'MsgBox Join(resuget, Chr(10))
'If numlig(Nolig) = 59 Then MsgBox chaine & Chr(10) & Join(resuget, Chr(10))
            If UBound(resuget) = 0 Then
                ' ERROR
                ' alimentation du CR ???
newLogLine = Array("", "ALERTE", FL0QU.NAME, "", "", "", chaine, (LigDeb0No + Nolig - 1), "Bouclage non résolu")
logNom = alimLog(logNom, newLogLine)
            Else
                nu = -1
                instance = listRes(Nolig)
                If InStr(Trim(listBoucle(Nolig)), "LAST") > 0 Then instance = resuget(UBound(resuget))
                If InStr(Trim(listBoucle(Nolig)), "FIRST") > 0 Then instance = resuget(LBound(resuget))
                ' cas ou une modalité = le EACH
                '''If InStr(Trim(listBOUCLE(Nolig)), "LAST") < 1 And InStr(Trim(listBOUCLE(Nolig)), "FIRST") < 1 Then
                    '''instance = chaineWithout & Trim(listBOUCLE(Nolig))
                    'if instance
                    'instance = Replace(d, f)
                    '''instanceList = getExtended("bou", True, instance, NOMENCLA, CInt(numlig(Nolig)), "QUANTITE", FLQUA.NAME, "L")
'''MsgBox chaine
                    '''If UBound(instanceList) = 0 Then
                        '''nu = -1
                    '''Else
                        '''For n = LBound(resuget) To UBound(resuget)
                            '''If instanceList(UBound(instanceList)) = resuget(n) Then
                                '''nu = n
                                '''Exit For
                            '''End If
                        '''Next
                    '''End If
                '''End If
                ''If nu = -1 Then
                    ' ERROR
                '''Else
                    If listRes(Nolig) = instance Then
                        ' calcul de la nouvelle équation
                        If Trim(listVALUE(Nolig)) <> "" Then
                            equatBoucle = getSolvedEquation(listVALUE(Nolig), FL0QU, NOMENCLA, areaList, scenarioList, numlig, listPer, listAre, listSce, listRes, listQua, Nolig)
                            valeurboucle = equatBoucle
                        Else
                            valeurboucle = "1"
                        End If
                        aj = "][" & listAre(Nolig) & "][" & listSce(Nolig) & "]." & listQua(Nolig) & "(t);"
                        resugetboucle = "[" & Join(resuget, aj & "[") & aj
                        adet = resugetboucle
                        '''resugetboucle = Replace(resugetboucle, "[" & resuget(nu) & aj, "")
                        resugetboucle = Replace(resugetboucle, "[" & instance & aj, "")
'MsgBox chaine & Chr(10) & valeurboucle & Chr(10) & adet & Chr(10) & Len(resugetboucle) & Chr(10) & Join(resuget, Chr(10))
                        If resugetboucle = "" Then
                            resugetboucle = "0"
                        Else
                            resugetboucle = valeurboucle & "-somme(" & Mid(resugetboucle, 1, Len(resugetboucle) - 1) & ")"
                        End If
                        listEqu(Nolig) = resugetboucle
                        listBoucle(Nolig) = "@" & listBoucle(Nolig)
                        suiteOK = False
                    End If
                '''End If
            End If
            End If
        End If
        If equat <> "" And suiteOK Then
            equat = getSolvedEquation(equat, FL0QU, NOMENCLA, areaList, scenarioList, numlig, listPer, listAre, listSce, listRes, listQua, Nolig)
            listEqu(Nolig) = equat
        End If
    Next
    ' Analyse syntaxique des équations
    Dim resAnaSynt() As String
    resAnaSynt = analyseSyntaxique(listEqu)
    Dim syntaxError As String
    Dim typSyntaxError As String
    Dim equSyntaxError As String
    For I = LBound(resAnaSynt) To UBound(resAnaSynt)
        If resAnaSynt(I) <> "" And resAnaSynt(I) <> "q" And InStr(listEqu(I), "ERROR") = 0 Then
            If syntaxError = "" Then
                syntaxError = syntaxError & "@" & numlig(I)
                typSyntaxError = typSyntaxError & "@" & resAnaSynt(I)
                equSyntaxError = equSyntaxError & "@" & listEqu(I)
            Else
                If InStr("@" & syntaxError & "@", "@" & numlig(I) & "@") = 0 Then
                    syntaxError = syntaxError & "@" & numlig(I)
                    typSyntaxError = typSyntaxError & "@" & resAnaSynt(I)
                    equSyntaxError = equSyntaxError & "@" & listEqu(I)
                End If
            End If
            listEqu(I) = "ERROR"
        End If
    Next
    If Len(syntaxError) > 0 Then
        splse = Split(syntaxError, "@")
        spltse = Split(typSyntaxError, "@")
        splese = Split(equSyntaxError, "@")
        For I = LBound(splse) To UBound(splse)
            If splse(I) <> "" Then
                ligne = Split("," & splse(I), ",")(UBound(Split("," & splse(I), ",")))
newLogLine = Array("", "ERREUR", FL0QU.NAME, "EQUATION", time(), "", spltse(I), ligne, "syntaxe incorrecte")
logNom = alimLog(logNom, newLogLine)
            End If
        Next
    End If
'MsgBox syntaxError & Chr(10) & typSyntaxError
    If ligSurcharge <> "" Then ligSurcharge = Mid(ligSurcharge, 2)
    If cibSurcharge <> "" Then cibSurcharge = Mid(cibSurcharge, 2)
    Dim aecrire() As String
    If UBound(listRes) > 0 Then
        ReDim aecrire(1 To UBound(listRes), 1 To 16)
        For Nolig = 1 To UBound(listRes)
            aecrire(Nolig, 1) = FL0QU.NAME
            aecrire(Nolig, 2) = numlig(Nolig)
            aecrire(Nolig, 3) = listRes(Nolig)
            aecrire(Nolig, 4) = listAre(Nolig)
            aecrire(Nolig, 5) = listSce(Nolig)
            aecrire(Nolig, 6) = listQua(Nolig)
            aecrire(Nolig, 7) = listTim(Nolig)
            aecrire(Nolig, 8) = listEqu(Nolig)
            aecrire(Nolig, 9) = listPer(Nolig)
            aecrire(Nolig, 10) = listNam(Nolig)
            aecrire(Nolig, 11) = listSCALE(Nolig)
            aecrire(Nolig, 12) = listUNIT(Nolig)
            aecrire(Nolig, 13) = listSUBSTITUTE(Nolig)
            aecrire(Nolig, 14) = listDEFAUT(Nolig)
            aecrire(Nolig, 15) = listVALUE(Nolig)
            aecrire(Nolig, 16) = listBoucle(Nolig)
        Next
    End If
    'DerLigSheet = FL0NO.Cells.SpecialCells(xlCellTypeLastCell).Row
    'For e = LigDeb0No To DerLigSheet
        'FL0NO.Cells(e, EXTENDED).Value = ""
    'Next
    'For e = 1 To UBound(extendNb)
        'FL0NO.Cells(LigDeb0No - 1 + newNumLig(e), EXTENDED).Value = extendNb(e)
    'Next
    ' Ecriture des résultats
    Dim scs As String
    If nbSurcharge > 1 Then scs = "s"
    If nbSurcharge < 2 Then scs = ""
'MsgBox "cibSurcharge=" & cibSurcharge
'MsgBox "ligSurcharge=" & ligSurcharge
'MsgBox "nbSurcharge=" & nbSurcharge
'cibSurcharge = "cibSurcharge"
    Dim lgs As String
    If nbLigOK > 1 Then lgs = "s"
    If nbLigOK < 2 Then lgs = ""
    Dim rts As String
    If nbRetraitsEnTout > 1 Then rts = "s"
    If nbRetraitsEnTout < 2 Then rts = ""
    Dim ets As String
    If UBound(listRes) > 1 Then ets = "s"
    If UBound(listRes) < 2 Then ets = ""
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", cibSurcharge, ligSurcharge, nbSurcharge & " surcharge" & scs)
logNom = alimLog(logNom, newLogLine)
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", cibRetrait, ligRetrait, nbRetraitsEnTout & " retrait" & rts)
logNom = alimLog(logNom, newLogLine)
newLogLine = Array("", "INFO", FL0QU.NAME, "Statistiques", time(), "", "", nbLigOK & " ligne" & lgs, UBound(listRes) & " quantité" & ets & " générée" & ets)
logNom = alimLog(logNom, newLogLine)
    Dim entete() As String
    ReDim entete(1 To 1, 1 To 17)
    entete(1, 1) = "Feuille"
    entete(1, 2) = "Ligne"
    entete(1, 3) = "ENTITE"
    entete(1, 4) = "AREA"
    entete(1, 5) = "SCENARIO"
    entete(1, 6) = "QUANTITE"
    entete(1, 7) = "TIME"
    entete(1, 8) = "EQUATION"
    entete(1, 9) = "PERIMETRE"
    entete(1, 10) = "NOM"
    entete(1, 11) = "SCALE"
    entete(1, 12) = "UNIT"
    entete(1, 13) = "SUBSTITUTE"
    entete(1, 14) = "DEFAUT"
    entete(1, 15) = "VALUE"
    entete(1, 16) = "BOUCLE"
    entete(1, 17) = "POSITION"

    FLQUA.Range("A1:Q1").VALUE = entete
    If UBound(listRes) > 0 Then FLQUA.Range("A2:P" & (UBound(listRes) + 1)).VALUE = aecrire
    FL0QU.Range("AD" & LigDeb0No & ":AD" & (LigDeb0No + UBound(listNb, 1) - 1)).VALUE = listNb
    sngChrono = Timer - sngChrono
logNom = ajoutNbQuaToError(logNom, listNb, LigDeb0No)
newLogLine = Array("QUANTITES", "FIN", FL0QU.NAME, "GENERATION", time(), "", "", "", (Int(1000 * sngChrono) / 1000) & " s")
logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCQU.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    FLCQU.Range("A1:K" & UBound(logNom, 1)).CurrentRegion.Borders.LineStyle = xlContinuous
    resColoriage = coloriage(LigDeb0No, FL0QU, FLCQU)
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C).VALUE = nbLigOK & ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 1).VALUE = UBound(listRes) & ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 2).VALUE = resColoriage & ""
    FLCTL.Cells(g_CONTROL_QUANTITES_GEN_CR_L, g_CONTROL_QUANTITES_GEN_CR_C + 3).VALUE = Round((Int(1000 * sngChrono) / 1000)) & ""
Application.StatusBar = False
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setQuantites"): Exit Sub
End Sub
Function getSolvedEquation(equat As String, FL0QU As Worksheet, NOMENCLA() As String, areaList() As String, scenarioList() As String, numlig() As String, listPer() As String, listAre() As String, listSce() As String, listRes() As String, listQua() As String, Nolig As Long) As String
    listEquations = analyseTerme(equat)
    Dim doubleatraiter As String
    Dim equationAtraiter As String
    Dim equationAtraiterEnCours As String
    Dim equationAtraiterInit As String
    Dim perimEt As String
    If UBound(listEquations) > 0 Then
    ' liste des termes de l'équation
    ' pour chaque terme on calcule l'extension
'MsgBox Join(listEquations, Chr(10))
        perimEt = "[" & listRes(Nolig) & "][" & listAre(Nolig) & "][" & listSce(Nolig) & "]"
        splResn = Split(listRes(Nolig), ".")
        For I = LBound(listEquations) To UBound(listEquations)
            ' résolution des raccourcis
            splEqua = Split(listEquations(I), ".")
            equationAtraiterEnCours = ""
            equationAtraiterInit = listEquations(I)
            If UBound(splEqua) > 0 Then
                For s = LBound(splEqua) To UBound(splEqua)
                    If s = LBound(splEqua) Then
                        jcts = ""
                    Else
                        jcts = "."
                    End If
                    If splEqua(s) Like "*>>*" Then
                        doubleatraiter = splEqua(s)
                        avt = ""
                        If Left(doubleatraiter, 1) = "[" Then
                            avt = "["
                            doubleatraiter = Mid(doubleatraiter, 2)
                        End If
                        apt = ""
                        If Right(doubleatraiter, 1) = "]" Then
                            apt = "]"
                            doubleatraiter = Mid(doubleatraiter, 1, Len(doubleatraiter) - 1)
                        End If
                        aajouter = avt & remplaceDoubleSup(doubleatraiter, NOMENCLA, "", 0) & apt
                        equationAtraiterEnCours = equationAtraiterEnCours & jcts & aajouter
                    Else
                        equationAtraiterEnCours = equationAtraiterEnCours & jcts & splEqua(s)
                    End If
                Next
                equationAtraiter = equationAtraiterEnCours
            Else
                equationAtraiter = listEquations(I)
            End If
            termeEtendu = getPerimeter(equationAtraiter, listPer(Nolig), perimEt, NOMENCLA, areaList, scenarioList, CInt(numlig(Nolig)), FL0QU, listRes, listAre, listSce, listQua)
'MsgBox equationAtraiter & Chr(10) & listPer(Nolig) & Chr(10) & termeEtendu
            If termeEtendu <> "" And Mid(termeEtendu, 1, 5) <> "ERROR" Then
                equat = Replace(equat, equationAtraiterInit, termeEtendu)
            Else
                If termeEtendu <> "" Then
                    equat = termeEtendu
                Else
                    equat = "ERROR"
                End If
                Exit For
            End If
        Next
    End If
    getSolvedEquation = equat
End Function
Function getTimes(tim As String, list() As String) As String()
    splTimEach = Split(tim, ">")
    splTimavEach = splTimEach(0)
    splTimapEach = splTimEach(1)
    splTim = Split(splTimavEach, "$")
    avt = splTim(0)
    apt = splTim(1)
    Dim getTimesf() As String
    ReDim getTimesf(0)
    For I = LBound(list) To UBound(list)
        splTimel = Split(list(I), "@")
        splTav = splTimel(0)
        splTap = splTimel(1)
        spltime = Split(splTav, "$")
        avte = spltime(0)
        apte = spltime(1)
        If (avt = avte Or avt = "") And apt = apte Then
            getTimes0 = Split(splTap, ",")
            ReDim getTimesf(1 To (UBound(getTimes0) + 1))
            For g = LBound(getTimes0) To UBound(getTimes0)
                getTimesf(g + 1) = splTimavEach & ">" & getTimes0(g)
            Next
            Exit For
        End If
    Next
    getTimes = getTimesf
End Function
Function remplaceDateByVal(form As String, times() As String) As String
    ' remplacement des dates par leurs valeurs
    Dim ok As String
    Dim apres As String
    Dim apresm As String
    Dim avant As String
    Dim avantt As String
    Dim typ As String
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Pattern = "\$[\.>a-zA-Z0-9]+"
    reg.Global = True
    Dim sep As String
    Dim MatchSubst As String
    If reg.test(form) Then
        Set matches = reg.Execute(form)
        If matches.Count <> 0 Then
            For Each Match In matches
                ' Détermination de la valeur de la date
                MatchSubst = ""
                If UBound(Split(Match, ".")) > 0 Then sep = "."
                If UBound(Split(Match, ">")) > 0 Then sep = ">"
                avant = Split(Match, sep)(0)
                typ = Split(avant, "$")(1)
                apres = Split(Match & sep, sep)(1)
                For t = LBound(times) To UBound(times)
                    avantt = Split(times(t), "@")(0)
                    typt = Split(avantt, "$")(1)
                    If typt = typ Then
                        listD = Split(Split(times(t), "@")(1), ",")
                        If LCase(apres) = "first" Then
                            MatchSubst = listD(LBound(listD))
                            Exit For
                        End If
                        If LCase(apres) = "last" Then
                            MatchSubst = listD(UBound(listD))
                            Exit For
                        End If
                        If LCase(apres) = "last" Then apresm = listD(UBound(listD))
                        For d = LBound(listD) To UBound(listD)
                            If listD(d) = apres Then
                                MatchSubst = listD(d)
                                Exit For
                            End If
                        Next
                        Exit For
                    End If
                Next
                If MatchSubst <> "" Then
                    form = Replace(form, Match, MatchSubst)
                End If
            Next
        End If
    End If
    remplaceDateByVal = form
End Function


Function dateInDate(occ As String, qua As String, timeslu() As String) As String
' Détermination du time de l'occurrence occ contexte dans le domaine qua de la quantité
    ' manque le cas ou qua = une date ??????
    dateInDate = ""
    Dim sep As String
    Dim avant As String
    Dim typ As String
    Dim cat As String
    Dim avantt As String
    Dim typt As String
    Dim catt As String
    sep = "/"
    If UBound(Split(occ, ".")) > 0 Then sep = "."
    If UBound(Split(occ, ">")) > 0 Then sep = ">"
    Dim adroite As String
    Dim gau As String
    Dim dro As String
    Dim ndate As String
    Dim splocc() As String
    Dim apres As String
    Dim apresq As String
    Dim apresq1 As String
    Dim apresq2 As String
    Dim apresm As String
    Dim aaa As Integer
    If occ Like "*-*" Then
        splocc = Split(sep & occ, sep)
        adroite = splocc(UBound(splocc))
        gau = Split(adroite, "-")(0)
        dro = Split(adroite, "-")(1)
        ndate = "" & (CInt(gau) - CInt(dro))
        occ = Split(occ, sep)(0) & sep & ndate
    End If
    If occ Like "*+*" Then
        splocc = Split(sep & occ, sep)
        adroite = splocc(UBound(splocc))
        gau = Split(adroite, "+")(0)
        dro = Split(adroite, "+")(1)
        ndate = "" & (CInt(gau) + CInt(dro))
        occ = Split(occ, sep)(0) & sep & ndate
    End If
    avant = Split(occ, sep)(0)
    apres = Split(occ & sep, sep)(1)
'MsgBox occ & "<>" & apres
    sepq = "/"
    If UBound(Split(qua, ".")) > 0 Then sepq = "."
    If UBound(Split(qua, ">")) > 0 Then sepq = ">"
    apresq = Split(qua & sepq, sepq)(1)
    If occ = qua Then
        typ = Split(Split(Split(qua & "$", "$")(1), ".")(0), ">")(0)
        cat = Split(qua, "$")(0)
        dateInDate = qua
        For t = LBound(timeslu) To UBound(timeslu)
            avantt = Split(timeslu(t), "@")(0)
            typt = Split(avantt, "$")(1)
            catt = Split(avantt, "$")(0)
            If (catt = cat Or cat = "") And typt = typ Then
                If UBound(Split(Split(timeslu(t), "@")(1), ",")) > 0 Then
                    listD = Split(Split(timeslu(t), "@")(1), ",")
                    dateInDate = occ & ">" & listD(LBound(listD))
                Else
                    dateInDate = occ & ">"
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    If Split(qua & "$", "$")(1) = "" Then
        typ = ""
    Else
        typ = Split(Split(Split(qua & "$", "$")(1), ".")(0), ">")(0)
    End If
    cat = Split(qua, "$")(0)
    dateInDate = ""
'MsgBox occ & ":::" & QUA & Chr(10) & sepq & ":::" & sep & Chr(10) & apresq & Chr(10) & apres
    Dim apresa As String
    For t = LBound(timeslu) To UBound(timeslu)
        avantt = Split(timeslu(t), "@")(0)
        typt = Split(avantt, "$")(1)
        catt = Split(avantt, "$")(0)
        If (catt = cat Or cat = "") And (typt = typ Or typ = "") Then
            listD = Split(Split(timeslu(t), "@")(1), ",")
            apresm = apresq
            apresa = apres
            If LCase(apresq) = "first" Then apresm = listD(LBound(listD))
            If LCase(apresq) = "last" Then apresm = listD(UBound(listD))
            If LCase(apres) = "first" Then apresa = listD(LBound(listD))
            If LCase(apres) = "last" Then apresa = listD(UBound(listD))
'MsgBox cat & ":" & typ & Chr(10) & apresq & Chr(10) & apresa
            For d = LBound(listD) To UBound(listD)
                If apresm = "" Then
                    If listD(d) = apresa Then
                        dateInDate = avantt & ">" & listD(d)
                        Exit For
                    End If
                Else
                    If (listD(d) = apresm) And (apresa = apresm) Then
                        dateInDate = avantt & ">" & listD(d)
                        Exit For
                    End If
                End If
            Next
            Exit For
        End If
    Next
End Function

Function getTimeMatch(tim As String, ctx As String, timeslu() As String) As String
    ' Détermination du TIME à partir de tim, du contexte et des temps
    Dim datectx As String
    datectx = Split(ctx & ">", ">")(1)
    Dim avant As String
    Dim typ As String
    Dim cat As String
    Dim avantt As String
    Dim typt As String
    Dim catt As String
    
    If tim Like "*.last*" Then
        avant = Split(tim, ".last")(0)
        typ = Split(avant, "$")(1)
        cat = Split(avant, "$")(0)
        getTimeMatch = ""
        For t = LBound(timeslu) To UBound(timeslu)
            avantt = Split(timeslu(t), "@")(0)
            typt = Split(avantt, "$")(1)
            catt = Split(avantt, "$")(0)
            If (catt = cat Or cat = "") And typt = typ Then
                listD = Split(Split(timeslu(t), "@")(1), ",")
                If datectx = listD(UBound(listD)) Then
                    getTimeMatch = ctx
                Else
                    getTimeMatch = ""
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    If tim Like "*.first*" Then
        avant = Split(tim, ".first")(0)
        typ = Split(avant, "$")(1)
        cat = Split(avant, "$")(0)
        getTimeMatch = ""
        For t = LBound(timeslu) To UBound(timeslu)
            avantt = Split(timeslu(t), "@")(0)
            typt = Split(avantt, "$")(1)
            catt = Split(avantt, "$")(0)
            If (catt = cat Or cat = "") And typt = typ Then
                listD = Split(Split(timeslu(t), "@")(1), ",")
                If datectx = listD(LBound(listD)) Then
                    getTimeMatch = ctx
                Else
                    getTimeMatch = ""
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    If tim Like "*>last*" Then
        avant = Split(tim, ">last")(0)
        typ = Split(avant, "$")(1)
        cat = Split(avant, "$")(0)
        getTimeMatch = ""
        For t = LBound(timeslu) To UBound(timeslu)
            avantt = Split(timeslu(t), "@")(0)
            typt = Split(avantt, "$")(1)
            catt = Split(avantt, "$")(0)
            If (catt = cat Or cat = "") And typt = typ Then
                listD = Split(Split(timeslu(t), "@")(1), ",")
                If datectx = listD(UBound(listD)) Then
                    getTimeMatch = ctx
                Else
                    getTimeMatch = ""
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    If tim Like "*>first*" Then
        avant = Split(tim, ">first")(0)
        typ = Split(avant, "$")(1)
        cat = Split(avant, "$")(0)
        getTimeMatch = ""
        For t = LBound(timeslu) To UBound(timeslu)
            avantt = Split(timeslu(t), "@")(0)
            typt = Split(avantt, "$")(1)
            catt = Split(avantt, "$")(0)
            If (catt = cat Or cat = "") And typt = typ Then
                listD = Split(Split(timeslu(t), "@")(1), ",")
                If datectx = listD(LBound(listD)) Then
                    getTimeMatch = ctx
                Else
                    getTimeMatch = ""
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    If tim = "" Then
        ' cas où tim est vide
        getTimeMatch = ctx
        Exit Function
    End If
    If tim = ctx Then
        ' cas où tim = le contexte
        getTimeMatch = ctx
        Exit Function
    End If
    If tim = Left(ctx, Len(tim)) Then
        ' cas où tim = le type du contexte
        getTimeMatch = ctx
        Exit Function
    End If
    ' Détermination de la ligne de tim dans timeslu
    If Split(tim & ">", ">")(1) <> "" Then
        ' cas où tim a une date
        getTimeMatch = tim
        Exit Function
    Else
        ' cas où tim n'a pas de date
        Dim ptim As Integer
        ptim = 0
        For I = LBound(timeslu) To UBound(timeslu)
            splt = Split(timeslu(I), "@")(0)
            If InStr(tim, splt) > 0 Then ptim = I
        Next
        If ptim = 0 Then
            ' cas où le type de tim n'est pas défini
            getTimeMatch = ""
            Exit Function
        End If
        If ptim > 0 Then
            ' cas où le type de tim est défini
            splt = Split(timeslu(ptim), "@")(1)
            spltd = Split(splt, ",")
            For d = LBound(spltd) To UBound(spltd)
                If datectx = spltd(d) Then
                    getTimeMatch = tim
                    Exit Function
                End If
            Next
            getTimeMatch = ""
        End If
    End If
End Function


Sub comboListNom()
    'MsgBox "ICI"
    
      'With Me.cboComboBox
            'cboComboBox.AddItem "Domaine Alexis Rouge"
            'cboComboBox.AddItem "Domaine du Grand Crès Blanc"
            'cboComboBox.AddItem "Domaine du Grand Crès Muscat"
            'cboComboBox.AddItem "Domaine du Grand Crès Rosé"
            'cboComboBox.AddItem "Domaine du Grand Crès Rouge"
      'End With
End Sub
Sub ComboBox_Create()
    ActiveSheet.cbTime.Clear
    ActiveSheet.cbArea.Clear
    ActiveSheet.cbScenario.Clear
    ActiveSheet.cbNom.Clear
    ActiveSheet.cbQua.Clear
    ActiveSheet.ListData.Clear
    ActiveSheet.ListInput.Clear
    For Each WS In ThisWorkbook.Worksheets
        firstcell = Trim(WS.Cells(1, 1).VALUE)
        If firstcell Like "*DATE*" Then
            ActiveSheet.cbTime.AddItem (WS.NAME)
        End If
        If firstcell Like "*AREA*" Then
            ActiveSheet.cbArea.AddItem (WS.NAME)
        End If
        If firstcell Like "*SCENARIO*" Then
            ActiveSheet.cbScenario.AddItem (WS.NAME)
        End If
        If firstcell Like "*NOMENCLATURE*" Then
            ActiveSheet.cbNom.AddItem (WS.NAME)
        End If
        If firstcell Like "*QUANTITE*" Then
            ActiveSheet.cbQua.AddItem (WS.NAME)
        End If
        If firstcell Like "*NOP_Col*" Then
            ActiveSheet.ListData.AddItem (WS.NAME)
        End If
        If firstcell Like "*INPUT*" Then
            ActiveSheet.ListInput.AddItem (WS.NAME)
        End If
    Next WS
    ActiveSheet.cbTime.ListIndex = 0
    ActiveSheet.cbArea.ListIndex = 0
    ActiveSheet.cbScenario.ListIndex = 0
    ActiveSheet.cbNom.ListIndex = 0
    ActiveSheet.cbQua.ListIndex = 0
    ActiveSheet.Cells(9, 2).VALUE = ""
    ActiveSheet.Cells(9, 3).VALUE = ""
    ActiveSheet.Cells(6, 5).VALUE = ""
    ActiveSheet.Cells(6, 6).VALUE = ""
    ActiveSheet.Cells(6, 7).VALUE = ""
    ActiveSheet.Cells(6, 8).VALUE = ""
    ActiveSheet.Cells(7, 5).VALUE = ""
    ActiveSheet.Cells(7, 6).VALUE = ""
    ActiveSheet.Cells(7, 7).VALUE = ""
    ActiveSheet.Cells(7, 8).VALUE = ""
    ActiveSheet.Cells(9, 5).VALUE = ""
    ActiveSheet.Cells(9, 6).VALUE = ""
    ActiveSheet.Cells(9, 7).VALUE = ""
    ActiveSheet.Cells(9, 8).VALUE = ""
    
    ActiveSheet.Cells(9, 10).VALUE = ""
    ActiveSheet.Cells(9, 11).VALUE = ""
    ActiveSheet.Cells(9, 12).VALUE = ""
    ActiveSheet.Cells(9, 13).VALUE = ""
    
    ActiveSheet.Cells(9, 15).VALUE = ""
    ActiveSheet.Cells(9, 16).VALUE = ""
    ActiveSheet.Cells(9, 17).VALUE = ""
    ActiveSheet.Cells(9, 18).VALUE = ""
End Sub

Public Function ItemExist(mCol As Collection, key As String) As Boolean
    Dim V As Variant
    On Error Resume Next
    V = mCol(key)
    If Err.Number = 450 Or Err.Number = 0 Then
        ItemExist = True
    Else
        ItemExist = False
    End If
End Function
Function getName(entite As String, listN() As String, listA() As String, listS() As String) As String
    entite = Replace(entite, ".NAME.1", ".NAME")
    entite = Replace(entite, ".NAME.2", ".NAME")
    entite = Replace(entite, ".NAME.3", ".NAME")
'If entite Like "*.NAME*" Then MsgBox entite
    If Left(entite, 2) = "[[" Then
        p = InStr(entite, "]")
        entite = "[" & Mid(entite, p + 1)
    End If
    splEntite = Split(entite, "[")
    Dim reconst As String
    reconst = ""
    Dim enCours As String
    Dim datee As String
    Dim list() As String
    Dim av As String
    getName = entite
    If UBound(splEntite) > 0 Then
        reconst = splEntite(0)
        For I = (LBound(splEntite) + 1) To UBound(splEntite)
            enCours = splEntite(I)
            avant = Split(enCours, "]")(0)
            '''If splEntite(i) Like "*$*" Then
            If avant Like "*$*" Then
                avant = Split(enCours, "]")(0)
                If UBound(Split(enCours, "]")) < 1 Then
                    Exit Function
                End If
                apres = Split(enCours, "]")(1)
                datee = Split(avant & ">", ">")(1)
                enCours = datee & apres
            Else
                splcp = Split(enCours, "].NAME")
                If UBound(splcp) > 0 Then
                    avant = Split(enCours, "].NAME")(0)
                    apres = Split(enCours, "].NAME")(1)
                    '''remplaceDoubleSup(ent, NOMENCLA, FLICI.NAME, 0)
                    datee = Split(avant & ">", ">")(1)
                    list = listN
                    If Left(avant, 2) = "a>" Then list = listA
                    If Left(avant, 2) = "s>" Then list = listS
                    If Left(avant, 2) <> "a>" And Left(avant, 2) <> "s>" Then
                        av = avant
                        avant = remplaceDoubleSup(av, listN, "", 0)
                    End If
                    For e = LBound(list) To UBound(list)
                        splList0 = Split(list(e) & "@", "@")(0)
                        splList1 = Split(list(e) & "@", "@")(1)
                        ' enlever les attributs ???
'If splList0 Like "*MENAGE*" Then MsgBox avant & Chr(10) & splList0
                        If splList0 = avant Then
'MsgBox encours & "::" & avant & "::" & splList1
'If Left(avant, 2) = "s>" Then MsgBox "getName " & Chr(10) & avant & Chr(10) & datee & Chr(10) & splList0 & Chr(10) & splList1
                            If splList1 <> "" Then
                                datee = Split(Split(splList1, ":")(0), ".")(0)
                            Else
                                datee = Split(Split(StrReverse(Split(StrReverse(avant), ">")(0)), ":")(0), ".")(0)
                            End If
                            Exit For
                        End If
                    Next
'If Left(avant, 2) = "s>" Then MsgBox "getName " & Chr(10) & avant & Chr(10) & datee & Chr(10) & Join(list, Chr(10))
'MsgBox entite & "===" & datee & Chr(10) & Join(list, Chr(10))
                    enCours = datee & apres
                Else
                
'If UBound(Split(encours, "]")) < 1 Then MsgBox entite
                    avant = Split(enCours, "]")(0)
                    apres = Split(enCours, "]")(1)
                    If Left(apres, 1) <> "." Then
                        If Left(avant, 2) <> "a>" And Left(avant, 2) <> "s>" Then
                            av = avant
                            avant = remplaceDoubleSup(av, listN, "", 0)
                        End If
                        datee = Split(Split(Split(">" & avant, ">")(UBound(Split(">" & avant, ">"))), ":")(0), ".")(0)
                        enCours = datee & apres
                    Else
                        reconst = "[" & reconst
'MsgBox reconst
                    End If
                End If
            End If
            reconst = reconst & enCours
'MsgBox "getName" & Chr(10) & entite & Chr(10) & encours & Chr(10) & reconst
        Next
        getName = reconst
    Else
        getName = entite
    End If
    'getName = entite
End Function
Sub setSheetSelectedFormula()
    Dim listFeuilles() As String
    ReDim listFeuilles(0)
    ActiveSheet.Cells(9, 15).VALUE = ""
    ActiveSheet.Cells(9, 16).VALUE = ""
    ActiveSheet.Cells(9, 17).VALUE = ""
    ActiveSheet.Cells(9, 18).VALUE = ""
    Dim FLCDG As Worksheet
    Set FLCDG = Worksheets("CRFORMULA")
    derlig = Split(FLCDG.UsedRange.Address, "$")(4)
    LigDebCNo = 2
    For Nolig = 1 To derlig
        If FLCDG.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = FLCDG.Range("A" & FLCDG.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCDG.Columns.Count).End(xlToLeft).Column
    FLCDG.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    For I = 0 To ActiveSheet.ListData.ListCount - 1
        If ActiveSheet.ListData.Selected(I) Then
            If UBound(listFeuilles) = 0 Then
                ReDim listFeuilles(1 To 1)
            Else
                ReDim Preserve listFeuilles(1 To UBound(listFeuilles) + 1)
            End If
            listFeuilles(UBound(listFeuilles)) = ActiveSheet.ListData.list(I)
            setFormula (Mid(listFeuilles(UBound(listFeuilles)), 2))
        Else
            If ActiveSheet.Cells(9, 15) = "" Then
                ActiveSheet.Cells(9, 15).VALUE = "."
                ActiveSheet.Cells(9, 16).VALUE = "."
                ActiveSheet.Cells(9, 17).VALUE = "."
                ActiveSheet.Cells(9, 18).VALUE = "."
            Else
                ActiveSheet.Cells(9, 15).VALUE = ActiveSheet.Cells(9, 15).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 16).VALUE = ActiveSheet.Cells(9, 16).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 17).VALUE = ActiveSheet.Cells(9, 17).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 18).VALUE = ActiveSheet.Cells(9, 18).VALUE & Chr(10) & "."
            End If
        End If
    Next I
End Sub
Sub setFormulaSel()
    Dim feuille As String
    Dim checkFeuille As Boolean
    Dim nbFeuilles As Integer
    Dim ind As Integer
    nbFeuilles = 0
    g_withoutFormula = g_WB_Extra.Worksheets(g_CONTROL).cbSansFormula.VALUE
    ' Remise à vide des CR
    ' ???
    ' Boucler sur les noms dans la feuille
    For I = 1 To 99
        feuille = Trim(ActiveSheet.Cells(g_CONTROL_DATA_L - 1 + I, g_CONTROL_DATA_C).VALUE)
        If feuille <> "" Then
            ind = I
            checkFeuille = getSheetSelected(ind)
            If checkFeuille Then
                nbFeuilles = nbFeuilles + 1
                Call DisableExcelSoft
                setFormula (feuille)
                'g_WB_Extra.Worksheets(g_CONTROL).Select
                Call EnableExcelSoft
                butInfo.Show
            End If
        End If
    Next I
    g_WB_Extra.Worksheets(g_CONTROL).Select
    Application.StatusBar = "Les formules ont été générées"
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille du modèle n'est sélectionnée"
    End If
End Sub
Function getMetaSheetSelected(I As Integer) As String
    Dim checkFeuille As String
    If I = 1 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta1.VALUE
    If I = 2 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta2.VALUE
    If I = 3 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta3.VALUE
    If I = 4 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta4.VALUE
    If I = 5 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta5.VALUE
    If I = 6 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta6.VALUE
    If I = 7 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta7.VALUE
    If I = 8 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta8.VALUE
    If I = 9 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta9.VALUE
    If I = 10 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta10.VALUE
    If I = 11 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta11.VALUE
    If I = 12 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta12.VALUE
    If I = 13 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta13.VALUE
    If I = 14 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta14.VALUE
    If I = 15 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta15.VALUE
    If I = 16 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta16.VALUE
    If I = 17 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta17.VALUE
    If I = 18 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta18.VALUE
    If I = 19 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta19.VALUE
    If I = 20 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelMeta20.VALUE
    getMetaSheetSelected = checkFeuille
End Function
Function getSheetSelected(I As Integer) As String
    Dim checkFeuille As String
    If I = 1 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData1.VALUE
    If I = 2 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData2.VALUE
    If I = 3 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData3.VALUE
    If I = 4 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData4.VALUE
    If I = 5 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData5.VALUE
    If I = 6 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData6.VALUE
    If I = 7 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData7.VALUE
    If I = 8 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData8.VALUE
    If I = 9 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData9.VALUE
    If I = 10 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData10.VALUE
    If I = 11 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData11.VALUE
    If I = 12 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData12.VALUE
    If I = 13 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData13.VALUE
    If I = 14 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData14.VALUE
    If I = 15 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData15.VALUE
    If I = 16 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData16.VALUE
    If I = 17 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData17.VALUE
    If I = 18 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData18.VALUE
    If I = 19 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData19.VALUE
    If I = 20 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData20.VALUE
    If I = 21 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData21.VALUE
    If I = 22 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData22.VALUE
    If I = 23 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData23.VALUE
    If I = 24 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData24.VALUE
    If I = 25 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData25.VALUE
    If I = 26 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData26.VALUE
    If I = 27 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData27.VALUE
    If I = 28 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData28.VALUE
    If I = 29 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData29.VALUE
    If I = 30 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData30.VALUE
    If I = 31 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData31.VALUE
    If I = 32 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData32.VALUE
    If I = 33 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData33.VALUE
    If I = 34 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData34.VALUE
    If I = 35 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData35.VALUE
    If I = 36 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData36.VALUE
    If I = 37 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData37.VALUE
    If I = 38 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData38.VALUE
    If I = 39 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData39.VALUE
    If I = 40 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData40.VALUE
    If I = 41 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData41.VALUE
    If I = 42 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData42.VALUE
    If I = 43 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData43.VALUE
    If I = 44 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData44.VALUE
    If I = 45 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData45.VALUE
    If I = 46 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData46.VALUE
    If I = 47 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData47.VALUE
    If I = 48 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData48.VALUE
    If I = 49 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData49.VALUE
    If I = 50 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData50.VALUE
    If I = 51 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData51.VALUE
    If I = 52 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData52.VALUE
    If I = 53 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData53.VALUE
    If I = 54 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData54.VALUE
    If I = 55 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData55.VALUE
    If I = 56 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData56.VALUE
    If I = 57 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData57.VALUE
    If I = 58 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData58.VALUE
    If I = 59 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData59.VALUE
    If I = 60 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData60.VALUE
    If I = 61 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData61.VALUE
    If I = 62 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData62.VALUE
    If I = 63 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData63.VALUE
    If I = 64 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData64.VALUE
    If I = 65 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData65.VALUE
    If I = 66 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData66.VALUE
    If I = 67 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData67.VALUE
    If I = 68 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData68.VALUE
    If I = 69 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData69.VALUE
    If I = 70 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData70.VALUE
    If I = 71 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData71.VALUE
    If I = 72 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData72.VALUE
    If I = 73 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData73.VALUE
    If I = 74 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData74.VALUE
    If I = 75 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData75.VALUE
    If I = 76 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData76.VALUE
    If I = 77 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData77.VALUE
    If I = 78 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData78.VALUE
    If I = 79 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData79.VALUE
    If I = 80 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData80.VALUE
    If I = 81 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData81.VALUE
    If I = 82 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData82.VALUE
    If I = 83 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData83.VALUE
    If I = 84 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData84.VALUE
    If I = 85 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData85.VALUE
    If I = 86 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData86.VALUE
    If I = 87 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData87.VALUE
    If I = 88 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData88.VALUE
    If I = 89 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData89.VALUE
    If I = 90 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData90.VALUE
    If I = 91 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData91.VALUE
    If I = 92 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData92.VALUE
    If I = 93 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData93.VALUE
    If I = 94 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData94.VALUE
    If I = 95 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData95.VALUE
    If I = 96 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData96.VALUE
    If I = 97 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData97.VALUE
    If I = 98 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData98.VALUE
    If I = 99 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelData99.VALUE
    getSheetSelected = checkFeuille
End Function
Sub setDataToSheetSel()
    Dim feuille As String
    Dim checkFeuille As Boolean
    Dim nbFeuilles As Integer
    Dim ind As Integer
    nbFeuilles = 0
    ' Constitution de la liste de synonymes
    Call setSynonymes
    ' Lecture du paramètre : à partir de la dernière valeur connue
    g_lastValue = ThisWorkbook.Worksheets(g_CONTROL).cbLastValue.VALUE
    ' Boucler sur les noms dans la feuille
    For I = 1 To 99
        feuille = Trim(ActiveSheet.Cells(g_CONTROL_DATA_L - 1 + I, g_CONTROL_DATA_C).VALUE)
        If feuille <> "" Then
            ind = I
            checkFeuille = getSheetSelected(ind)
            If checkFeuille Then
                nbFeuilles = nbFeuilles + 1
                Call DisableExcelSoft
                setDataToSheet (feuille)
                g_WB_Extra.Worksheets(g_CONTROL).Activate
                g_WB_Extra.Worksheets(g_CONTROL).Select
                Call EnableExcelSoft
                butInfo.Show
            End If
        End If
    Next I
    g_WB_Extra.Worksheets(g_CONTROL).Select
    Application.StatusBar = "Les feuilles ont été alimentées"
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille du modèle n'est sélectionnée"
    End If
End Sub

Function normalisationString(str As String) As String
    str = normalisationStringCasse(str)
    str = normalisationStringOther(str)
    normalisationString = str
    resb = Filter(g_Synonymes, "," & str & ",", True)
    If UBound(resb) >= 0 Then
        normalisationString = Split(resb(LBound(resb)), ",")(0)
    End If
    normalisationString = Replace(normalisationString, "  ", " ")
End Function
Function getDerCol(FL) As Integer
    spl = Split(FL.UsedRange.Address, "$")
    If UBound(spl) > 3 Then
        getDerCol = Columns(spl(3)).Column
    Else
        getDerCol = Cells(1, FL.Columns.Count).End(xlToLeft).Column
    End If
End Function
Function getDerLig(FL) As Integer
    spl = Split(FL.UsedRange.Address, "$")
    If UBound(spl) > 3 Then
        getDerLig = spl(4)
    Else
        getDerLig = FL.Range("A" & FL.Rows.Count).End(xlUp).Row
    End If
End Function
Sub setDataToSheet(sheettoAlim As String)
    On Error GoTo errorHandler
    sngChrono = Timer
    openFileIfNot ("MODELE")
    openFileIfNot ("SOURCE")
    openFileIfNot ("TARGET")
    Dim nameTime As String
    nameTime = ActiveSheet.cbTime.VALUE
    Dim times As String
    times = Split(lectureTime(nameTime)(1), "@")(1)
    Dim listTimes() As String
    listTimes = Split(times, ",")
    Dim FLICI As Worksheet
    Dim FLINP As Worksheet
    Dim FLCTL As Worksheet
    Dim FLCDG As Worksheet
    targetName = getNameIfExists(getNameFromModel(sheettoAlim, "CIBLE"), g_WB_Target)
    If targetName = "" Then GoTo errorHandlerWsNotExists
    sourceNAme = getNameIfExists(getNameFromModel(sheettoAlim, "SOURCE"), g_WB_Source)
    If sourceNAme = "" Then Exit Sub
    Set FLCDG = setWS("CRIMPORT", g_WB_Modele)
    Set FLCTL = Worksheets(g_CONTROL)
    derlig = Split(FLCDG.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLCDG.Cells(Nolig, 1) <> "" Then LigDebCNo = Nolig + 1
    Next
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = getDerLig(FLCDG)
    DerColSheet = getDerCol(FLCDG)
    With FLCDG.Range("a1:" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Set FLICI = g_WB_Target.Worksheets(targetName)
    Dim newLogLine() As Variant
newLogLine = Array("DATAIMP", "DEBUT", FLICI.NAME, "IMPORTATION", time())
logNom = alimLog(logNom, newLogLine)
    Set FLINP = g_WB_Source.Worksheets(sourceNAme)
    ' boucle sur les lignes du fichier à alimenter
    Dim Cla() As Variant
    derligsheet = getDerLig(FLICI)
    DerColSheet = getDerCol(FLICI)
'MsgBox FLICI.NAME & ":" & DerLigSheet
    With FLICI.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
        ReDim Cla(1 To derligsheet, 1 To DerColSheet)
        Cla = .VALUE
        Cla = setDataTransport(Cla)
    End With
    Dim ClaC1() As Variant
    With FLICI.Range("a1:a" & derligsheet)
        ReDim ClaC1(1 To derligsheet, 1 To 1)
        ClaC1 = .VALUE
    End With
    Dim Sui() As Variant
    derligsheet = getDerLig(FLINP)
    DerColSheet = getDerCol(FLINP)
    Dim suiNew() As Variant
    Dim SuiOld() As Variant
    Dim SuiInt() As Variant
    With FLINP.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
        SuiOld = .VALUE
        SuiInt = setDataTertiaire(SuiOld)
        Sui = setDataTransport(SuiInt)
    End With
    Dim SuiC1() As Variant
    ReDim SuiC1(1 To UBound(Sui, 1), 1 To 1)
    For Nolig = 1 To UBound(Sui, 1)
    Next
    Dim cellule As String
    Dim suite As Integer
    suite = 0
    For Nolig = 1 To UBound(Sui, 1)
        SuiC1(Nolig, 1) = Sui(Nolig, 1)
    Next
    Dim nbLigImp As Integer
    nbLigImp = 0
    ' Match des contextes
    Dim ctc() As String
    ReDim ctc(1 To UBound(ClaC1, 1))
    Dim tic() As String
    ReDim tic(1 To UBound(Cla, 2))
    Dim cts() As String
    ReDim cts(1 To UBound(SuiC1, 1))
    Dim tis() As String
    ReDim tis(1 To UBound(Sui, 2))
    Dim tisn() As String
    ReDim tisn(1 To UBound(Sui, 2))
    For I = LBound(ClaC1, 1) To UBound(ClaC1, 1)
        ctc(I) = ClaC1(I, 1)
    Next
    For I = LBound(SuiC1, 1) To UBound(SuiC1, 1)
        cts(I) = SuiC1(I, 1)
    Next
    Dim ctcok() As Integer
    ReDim ctcok(1 To UBound(ClaC1, 1))
    Dim ctsok() As Integer
    ReDim ctsok(1 To UBound(SuiC1, 1))
    For j = LBound(ctsok) To UBound(ctsok)
        ctsok(j) = 0
    Next
    Dim reprise As Integer
    reprise = LBound(cts)
    Dim rep As Integer
    rep = reprise
    Dim nbok As Integer
    nbok = 0
    Dim nbnb As Integer
    nbnb = 0
    Dim distance As Integer
    distance = 10
    Dim ccc As String
    Dim sss As String
    Dim inclus As Integer
    Dim concat As String
    For I = LBound(ctc) To UBound(ctc)
        reprise = rep
        ctcok(I) = 0
        concat = ""
        If Trim(ctc(I)) <> "" Then
            nbnb = nbnb + 1
            inclus = 0
            For j = WorksheetFunction.Max(reprise - distance, LBound(cts)) To WorksheetFunction.Min(reprise + distance, UBound(cts))
                If Trim(cts(j)) <> "" Then
                    ccc = normalisationStringCasse(ctc(I))
                    sss = normalisationStringCasse(cts(j))
                    If ctsok(j) = 0 Then
                        ccc1 = normalisationString(ccc)
                        sss1 = normalisationString(sss)
                        If (InStr(sss1, ccc1) > 0 Or InStr(ccc1, sss1) > 0) And inclus = 0 Then
                            ' on retient si inclusion
                            inclus = j
                        End If
                        If sss = ccc Then
                            ' si il n'y a pas d'inclusion ou si l'inclusion est plus éloignée que l'égalité
                            If inclus = 0 Or Math.Abs(inclus - I) >= (Math.Abs(j - I) - 1) Then
                                rep = WorksheetFunction.Min(j + 1, UBound(cts))
                                ctcok(I) = j
                                ctsok(j) = I
                                nbok = nbok + 1
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
            If ctcok(I) = 0 Then
                If inclus > 0 Then
                    ctcok(I) = inclus
                    ctsok(inclus) = I
                    nbok = nbok + 1
                End If
            End If
        End If
    Next
    nbLigImp = nbok
    Dim aimp As Integer
    Dim aimplist As String
    aimplist = ""
    Dim aimpn As Integer
    aimpn = 0
    Dim aremp As Integer
    aremp = 0
    Dim arempn As Integer
    arempn = 0
    Dim cellulebis As String
    Dim celluleimp As String
    Dim nlc As Integer
    Dim nls As Integer
    nlc = 0
    nls = 0
    Dim NoColDate As Integer
    Dim vals As String
    For Nolig = 1 To UBound(Cla, 1)
'MsgBox Nolig & ":" & Cla(Nolig, 1)
        aimp = 0
        aremp = 0
        cellule = Replace(Cla(Nolig, 1), "#", "")
        Cla(Nolig, 1) = cellule
        If Nolig <= UBound(Sui, 1) Then
            For NoCol = 2 To UBound(Sui, 2)
                celluleimp = Sui(Nolig, NoCol)
'If Nolig = 9 And NoCol > 11 Then MsgBox NoCol & ":" & celluleimp & ":" & InStr("," & times & ",", "," & celluleimp & ",")
                If InStr("," & times & ",", "," & celluleimp & ",") > 0 Then
                    If nls <> Nolig Then
                        ' NON on ne remet pas à vide car ne marche pas avec HM!!!
                        'For i = LBound(tis) To UBound(tis)
                            'tisn(i) = ""
                        'Next
                    End If
                    'If Nolig = 9 Then MsgBox NoCol & ":" & celluleimp
                    tis(NoCol) = celluleimp
                    tisn(NoCol) = "1"
                    nls = Nolig
                End If
            Next
'If Nolig = 358 Or Nolig = 357 Then MsgBox Join(tis, ",")
            If nls = Nolig Then
                For NoCol = 2 To UBound(Sui, 2)
                    tisn(NoCol) = ""
                    celluleimp = Sui(Nolig, NoCol)
                    If InStr("," & times & ",", "," & celluleimp & ",") > 0 Then
                        tisn(NoCol) = "1"
                    End If
                Next
            End If
        End If
'If Nolig = 9 Then MsgBox Nolig & Chr(10) & Join(tis, ",")
        If Nolig <= UBound(Cla, 1) Then
            For NoCol = 2 To UBound(Cla, 2)
                celluleimp = Cla(Nolig, NoCol)
                If InStr("," & times & ",", "," & celluleimp & ",") > 0 Then
                    If nlc <> Nolig Then
                        For I = LBound(tic) To UBound(tic)
                            tic(I) = ""
                        Next
                    End If
                    tic(NoCol) = celluleimp
                    nlc = Nolig
                End If
            Next
        End If
'MsgBox Nolig & Chr(10) & Join(tic, ",") & Chr(10) & Join(tis, ",") & Chr(10) & Join(tisn, ",")
        For NoCol = 2 To UBound(Cla, 2)
            cellulebis = ""
            On Error Resume Next
            cellulebis = Cla(Nolig, NoCol)
            cellule = Replace(cellulebis, "#", "")
            Cla(Nolig, NoCol) = cellule
            If Left(cellule, 2) = "I@" Or Left(cellule, 2) = "S@" Then
                aremp = Nolig
'If Nolig = 27 And NoCol = 18 Then MsgBox Cla(Nolig, NoCol) & ":" & Cla(Nolig, NoCol - 1)
'If Nolig = 27 And (NoCol = 24 Or NoCol = 30) Then MsgBox NoCol & "::" & Nolig & "::" & ctcok(Nolig) & Chr(10) & tic(NoCol) & Chr(10) & Join(tis, ",")
                If ctcok(Nolig) > 0 And InStr(Join(tis, ","), tic(NoCol)) > 0 Then
                    NoColDate = -1

                    If NoCol <= UBound(tis) Then
                        vals = tis(NoCol)
                    Else
                        vals = "XXX"
                    End If
                    If tic(NoCol) = vals And tisn(NoCol) = "1" Then
'If Nolig = 27 And NoCol = 18 Then MsgBox vals
                        'NoColDate = NoCol
                        'nc = getPosInList(tic, CInt(NoCol))
                        'ns = getPosInList(tis, CInt(NoCol))
                        For I = LBound(tis) To UBound(tis)
                            NoColDate = -1
                            nc = getPosInList(tic, CInt(NoCol))
                            If tic(NoCol) = tis(I) Then
                                NoColDate = I
                                If nc = getPosInList(tis, CInt(I)) Then Exit For
                            End If
                        Next
'If Nolig = 27 And (NoCol = 24 Or NoCol = 30) Then MsgBox NoCol & "::" & NoColDate
                    Else
'If Nolig = 5 And NoCol = 2 Then MsgBox Join(tic, ",") & Chr(10) & Join(tis, ",") & Chr(10) & Join(tisn, ",")
                        For I = LBound(tis) To UBound(tis)
                            NoColDate = -1
                            nc = getPosInList(tic, CInt(NoCol))
                            If tic(NoCol) = tis(I) And tisn(I) = "1" Then
                                NoColDate = I
'If Nolig = 27 And NoCol = 24 Then MsgBox Cla(Nolig, NoCol)
                                If nc = getPosInList(tis, CInt(I)) Then Exit For
                            End If
                        Next
                    End If
'If Nolig = 27 And (NoCol = 24 Or NoCol = 30) Then MsgBox NoCol & "::" & NoColDate & ":" & Sui(ctcok(Nolig), NoColDate)
                    If NoColDate = -1 Then
                        If tic(NoCol) = vals Then
                            NoColDate = NoCol
                        Else
                            For I = LBound(tis) To UBound(tis)
                                NoColDate = -1
                                nc = getPosInList(tic, CInt(NoCol))
                                If tic(NoCol) = tis(I) Then
                                    NoColDate = I
                                    If nc = getPosInList(tis, CInt(I)) Then Exit For
                                End If
                            Next
                        End If
                    End If
                    If NoColDate > -1 Then
                        If Sui(ctcok(Nolig), NoColDate) = "" Then
                            ' Cas où lavaleur est vide ==> rechercher de la date antérieure
                            'MsgBox Nolig & ":" & NoCol & Chr(10) & tis(NoColDate) & Chr(10) & Join(tis, ",")
                            For I = LBound(tis) To NoColDate
                                If Sui(ctcok(Nolig), I) <> "" And tis(I) <> "" Then NoColDate = I
                            Next
                        End If
'MsgBox Nolig & ":" & NoCol & "=" & tic(NoCol) & "=" & tis(NoColDate) & "::" & NoColDate
                        Cla(Nolig, NoCol) = Sui(ctcok(Nolig), NoColDate)
                    End If
                Else
                    If ctcok(Nolig) < 1 Then
                        aimp = Nolig
                    Else
                        If g_lastValue Then
                            If tic(NoCol) <> "" Then
                                ' Repérage de la date la plus proche inférieure
                                diffMin = CInt(tic(NoCol))
                                NoColDate = -1
                                For I = LBound(tis) To UBound(tis)
                                    If tis(I) <> "" Then
                                        If CInt(tis(I)) < CInt(tic(NoCol)) Then
                                            diff = CInt(tic(NoCol)) - CInt(tis(I))
                                            If diff < diffMin Then
                                                diffMin = diff
                                                NoColDate = I
                                            End If
                                        End If
                                    End If
                                Next
                                If NoColDate > -1 Then
                                    If Trim(Sui(ctcok(Nolig), NoColDate)) = "" Then
'If Nolig = 358 Or Nolig = 357 Then MsgBox Nolig & ":" & NoCol & ":tic(NoCol)=" & tic(NoCol) & ":=" & NoColDate & ":" & LBound(tis) & Chr(10) & Join(tis, ",")
                                        ' Cas où la valeur est vide ==> rechercher de la date antérieure
                                        For I = LBound(tis) To UBound(tis)
    'MsgBox Nolig & ":" & NoCol & ":" & tic(NoCol) & "<" & i & "<" & NoColDate & ":" & LBound(tis) & Chr(10) & tis(i)
'If (Nolig = 358 Or Nolig = 357) And i < 4 Then MsgBox Nolig & ":" & NoCol & ":" & ctcok(Nolig) & ":" & i & Chr(10) & Trim(Sui(ctcok(Nolig), i)) & Chr(10) & CInt(tis(i)) & "<=" & CInt(tic(NoCol))
                                            If Trim(Sui(ctcok(Nolig), I)) <> "" And tis(I) <> "" Then
'If Nolig = 358 Then MsgBox Nolig & ":" & NoCol & ":" & i & Chr(10) & CInt(tis(i)) & "<=" & CInt(tic(NoCol))
                                                If CInt(tis(I)) <= CInt(tic(NoCol)) Then
                                                    NoColDate = I
'If Nolig = 358 Then MsgBox Nolig & ":" & NoCol & ":" & NoColDate
                                                End If
    'MsgBox Nolig & ":" & NoCol & ":" & tic(NoCol) & ":" & NoColDate
                                            End If
                                        Next
                                    End If
                                    Cla(Nolig, NoCol) = Sui(ctcok(Nolig), NoColDate)
                                Else
                                    aimp = Nolig
                                End If
                            Else
                                aimp = Nolig
                            End If
                        Else
                            aimp = Nolig
                        End If
                    End If
                End If
            End If
        Next
        If aremp > 0 Then
            arempn = arempn + 1
        End If
        If aimp > 0 Then
            aimplist = aimplist & ";" & aimp
            aimpn = aimpn + 1
        End If
    Next
'MsgBox Join(tic, ",") & Chr(10) & Join(tis, ",") & Chr(10) & Join(tisn, ",")
    derligsheet = getDerLig(FLICI)
    DerColSheet = getDerCol(FLICI)
'MsgBox Cla(27, 24) & ":" & Cla(27, 30)
    With FLICI.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
        .VALUE = Cla
    End With
    If aimplist <> "" Then
'MsgBox aimplist
        If Len(aimplist) > 50 Then
            lll = ""
        Else
            c = aimplist
        End If
newLogLine = Array("", "ERREUR", FLICI.NAME, "IMPORTATION", time(), "", "", Mid(lll, 2), aimpn & " lignes non remplies")
logNom = alimLog(logNom, newLogLine)
    End If
newLogLine = Array("", "INFO", FLICI.NAME, "IMPORTATION", time(), "", "", "", nbLigImp & " lignes importées")
logNom = alimLog(logNom, newLogLine)
fin:
    sngChrono = Timer - sngChrono
newLogLine = Array("DATAIMP", "FIN", FLICI.NAME, "IMPORTATION", time(), "", "", "", (Int(1000 * sngChrono) / 1000) & " s")
logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCDG.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    FLCDG.Range("A1:K" & UBound(logNom, 1)).CurrentRegion.Borders.LineStyle = xlContinuous
    l = getLineFrom(sheettoAlim, FLCTL, g_CONTROL_DATA_L, g_CONTROL_DATA_C)
    FLCTL.Cells(l, g_CONTROL_DATA_DAT_C).VALUE = "" '''''& nbLigImp
    FLCTL.Cells(l, g_CONTROL_DATA_DAT_C + 1).VALUE = "" ''''''& arempn
    FLCTL.Cells(l, g_CONTROL_DATA_DAT_C + 2).VALUE = "" & aimpn
    FLCTL.Cells(l, g_CONTROL_DATA_DAT_C + 3).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setDataToSheet"): Exit Sub
errorHandlerWsNotExists: Call onErrDo("La feuille n'existe pas", "setDataToSheet"): Exit Sub
End Sub
Function getPosInList(list() As String, pos As Long) As Long
    getPosInList = 0
    For I = LBound(list) To pos
        If list(I) = list(pos) Then getPosInList = getPosInList + 1
    Next
End Function

Function setDataTertiaire(inp() As Variant) As Variant
    setDataTertiaire = inp
    Dim gap As Integer
    Dim gapPlus As Integer
    Dim gapBool As Boolean
    gap = 0
    gapPlus = 0
    gapBool = True
    'FEUILLE DE SAISIE DES HYPOTHESES TRANSPORT
    If inp(1, 1) = "FEUILLE DE SAISIE DES HYPOTHESES TERTIAIRE" Then
        If InStr(inp(5, 7), "Demande") > 0 Then
'MsgBox LBound(setDataTertiaire, 1) & ":" & UBound(setDataTertiaire, 1) & Chr(10) & LBound(setDataTertiaire, 2) & ":" & UBound(setDataTertiaire, 2)
            Dim res() As Variant
            Dim dates() As Integer
            ReDim res(1 To UBound(inp, 1) + 66, 1 To UBound(inp, 2))
            ReDim dates(1 To UBound(res, 2))
            For I = LBound(res, 1) To UBound(res, 1)
                If I = 57 Then gap = 0
                If I = 90 Then gap = 0
                If I > 56 And I < 90 Then
                    If (inp(I - 52 + gap, 7) & inp(I - 52 + gap, 8)) = "" Then
                        gapPlus = 0
                        For k = (I - 52 + gap + 1) To (I - 52 + gap + 6)
                            gapPlus = gapPlus + 1
                            If (inp(k, 7) & inp(k, 8)) <> "" Then
                                Exit For
                            End If
                        Next
                        gap = gap + gapPlus
                    End If
                End If
                If I > 89 And I < 122 Then
                    If (inp(I - 85 + gap, 13) & inp(I - 85 + gap, 14)) = "" Then
                        gapPlus = 0
                        For k = (I - 85 + gap + 1) To (I - 85 + gap + 6)
                            gapPlus = gapPlus + 1
                            If (inp(k, 13) & inp(k, 14)) <> "" Then
                                Exit For
                            End If
                        Next
                        gap = gap + gapPlus
                    End If
                End If
                For j = LBound(res, 2) To UBound(res, 2)
                    If I < 4 Then
                        res(I, j) = inp(I, j)
                    End If
                    If I > 122 Then
                        If (I - 67) <= UBound(inp, 1) Then
                            res(I, j) = inp(I - 67, j)
                        End If
                    End If
                    If I > 3 And I <= 56 Then
                        If j < 4 Then res(I, j) = inp(I, j)
                    End If
                    
                    If I > 56 And I < 90 Then
                        If j < 4 Then res(I, j) = inp(I - 52 + gap, j + 6)
                    End If
                    If I > 89 And I < 122 Then
                        If j < 4 Then res(I, j) = inp(I - 85 + gap, j + 12)
                    End If
                Next
            Next
            For I = LBound(res, 1) To UBound(res, 1)
                For j = LBound(res, 2) To UBound(res, 2)
                    If IsNumeric(res(I, j)) Then
                        If CInt(res(I, j)) > 2000 And CInt(res(I, j)) < 2100 Then
'If i = 300 And j = 3 Then MsgBox res(i, 1) & ":" & res(i, 2) & ":" & res(i, 3) & Chr(10) & dates(j)
                            If dates(j) = 0 Then
                                dates(j) = CInt(res(I, j))
                            Else
                                If dates(j) <> CInt(res(I, j)) Then
                                    res(I, j) = "" & dates(j)
                                End If
                            End If
                        End If
                    End If
                Next
            Next
            'MsgBox dates(1) & ":" & dates(2) & ":" & dates(3)
            setDataTertiaire = res
'Dim FLICI As Worksheet
'Set FLICI = g_WB_Target.Worksheets("HTtest")
'With FLICI.Range("a1:" & DecAlph(UBound(res, 2)) & UBound(res, 1))
        '.VALUE = res
'End With
        End If
    End If
End Function
Function setDataTransport(inp() As Variant) As Variant
    setDataTransport = inp
    Dim gap As Integer
    Dim gapPlus As Integer
    Dim gapBool As Boolean
    gap = 0
    gapPlus = 0
    gapBool = True
    If inp(1, 1) = "FEUILLE DE SAISIE DES HYPOTHESES TRANSPORT" Then
        For I = LBound(inp, 1) To UBound(inp, 1)
            If LCase(inp(I, 1)) = LCase("Evolution CU (base 100)") Then
                inp(I, 1) = inp(I, 1) & " " & Split(inp(I - 1, 1), " ")(0)
            End If
        Next
        setDataTransport = inp
    End If
End Function
Sub initializeForSheetGeneration()
    Call DisableExcel
    openFileIfNot ("MODELE")
    Dim FLCDG As Worksheet
    Set FLCDG = g_WB_Modele.Worksheets("CRDATA")
    derlig = Split(FLCDG.UsedRange.Address, "$")(4)
    LigDebCNo = 2
    For Nolig = 1 To derlig
        If FLCDG.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = FLCDG.Range("A" & FLCDG.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCDG.Columns.Count).End(xlToLeft).Column
    FLCDG.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    Dim FLCIG As Worksheet
    Set FLCIG = g_WB_Modele.Worksheets("CRIMPORT")
    derlig = Split(FLCIG.UsedRange.Address, "$")(4)
    LigDebCNo = 2
    For Nolig = 1 To derlig
        If FLCIG.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    derligsheet = FLCIG.Range("A" & FLCIG.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCIG.Columns.Count).End(xlToLeft).Column
    FLCIG.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    Dim FLCFG As Worksheet
    Set FLCFG = g_WB_Modele.Worksheets("CRFORMULA")
    derlig = Split(FLCFG.UsedRange.Address, "$")(4)
    LigDebCNo = 2
    For Nolig = 1 To derlig
        If FLCFG.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    derligsheet = FLCFG.Range("A" & FLCFG.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCFG.Columns.Count).End(xlToLeft).Column
    FLCFG.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    ' on clear les positions dans la feuille QUANTITY
    Dim FLQUA As Worksheet
    Set FLQUA = g_WB_Modele.Worksheets("QUANTITY")
    FLQUA.Range("Q:Q").Clear
    FLQUA.Cells(1, 17).VALUE = "POSITION"
    ' mise à vide des CR
    For ro = 1 To 20
        For co = 1 To 19
        ThisWorkbook.Worksheets(g_CONTROL).Cells(g_CONTROL_DATA_L - 1 + ro, g_CONTROL_DATA_C + co).VALUE = ""
        Next
    Next
    Call EnableExcel
End Sub
Sub setSheetSelected()
    Dim listFeuilles() As String
    ReDim listFeuilles(0)
    ActiveSheet.Cells(9, 5).VALUE = ""
    ActiveSheet.Cells(9, 6).VALUE = ""
    ActiveSheet.Cells(9, 7).VALUE = ""
    ActiveSheet.Cells(9, 8).VALUE = ""
    For I = 0 To ActiveSheet.ListData.ListCount - 1
        If ActiveSheet.ListData.Selected(I) Then
            If UBound(listFeuilles) = 0 Then
                ReDim listFeuilles(1 To 1)
            Else
                ReDim Preserve listFeuilles(1 To UBound(listFeuilles) + 1)
            End If
            listFeuilles(UBound(listFeuilles)) = ActiveSheet.ListData.list(I)
            setSheet (listFeuilles(UBound(listFeuilles)))
        Else
            If ActiveSheet.Cells(9, 5) = "" Then
                ActiveSheet.Cells(9, 5).VALUE = "."
                ActiveSheet.Cells(9, 6).VALUE = "."
                ActiveSheet.Cells(9, 7).VALUE = "."
                ActiveSheet.Cells(9, 8).VALUE = "."
            Else
                ActiveSheet.Cells(9, 5).VALUE = ActiveSheet.Cells(9, 5).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 6).VALUE = ActiveSheet.Cells(9, 6).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 7).VALUE = ActiveSheet.Cells(9, 7).VALUE & Chr(10) & "."
                ActiveSheet.Cells(9, 8).VALUE = ActiveSheet.Cells(9, 8).VALUE & Chr(10) & "."
            End If
        End If
    Next I
End Sub
Sub setInputSelected()
    Dim listFeuilles() As String
    ReDim listFeuilles(0)
    ActiveSheet.Cells(11, 5).VALUE = ""
    ActiveSheet.Cells(11, 6).VALUE = ""
    ActiveSheet.Cells(11, 7).VALUE = ""
    ActiveSheet.Cells(11, 8).VALUE = ""
    For I = 0 To ActiveSheet.ListInput.ListCount - 1
        If ActiveSheet.ListInput.Selected(I) Then
            If UBound(listFeuilles) = 0 Then
                ReDim listFeuilles(1 To 1)
            Else
                ReDim Preserve listFeuilles(1 To UBound(listFeuilles) + 1)
            End If
            listFeuilles(UBound(listFeuilles)) = ActiveSheet.ListInput.list(I)
            setInput (listFeuilles(UBound(listFeuilles)))
        Else
            If ActiveSheet.Cells(11, 5) = "" Then
                ActiveSheet.Cells(11, 5).VALUE = "."
                ActiveSheet.Cells(11, 6).VALUE = "."
                ActiveSheet.Cells(11, 7).VALUE = "."
                ActiveSheet.Cells(11, 8).VALUE = "."
            Else
                ActiveSheet.Cells(11, 5).VALUE = ActiveSheet.Cells(11, 5).VALUE & Chr(10) & "."
                ActiveSheet.Cells(11, 6).VALUE = ActiveSheet.Cells(11, 6).VALUE & Chr(10) & "."
                ActiveSheet.Cells(11, 7).VALUE = ActiveSheet.Cells(11, 7).VALUE & Chr(10) & "."
                ActiveSheet.Cells(11, 8).VALUE = ActiveSheet.Cells(11, 8).VALUE & Chr(10) & "."
            End If
        End If
    Next I
End Sub
Sub setInputSelectedFrom()
    Dim feuille As String
    Dim checkFeuille As Boolean
    Dim nbFeuilles As Integer
    
    nbFeuilles = 0
    ' Remise à vide des CR
    ' ???
    ' Boucler sur les noms dans la feuille
    For I = 1 To 99
        feuille = Trim(ActiveSheet.Cells(g_CONTROL_INPUT_L - 1 + I, g_CONTROL_INPUT_C).VALUE)
        If feuille <> "" Then
            If I = 1 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput1.VALUE
            If I = 2 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput2.VALUE
            If I = 3 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput3.VALUE
            If I = 4 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput4.VALUE
            If I = 5 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput5.VALUE
            If I = 6 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput6.VALUE
            If I = 7 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput7.VALUE
            If I = 8 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput8.VALUE
            If I = 9 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput9.VALUE
            If I = 10 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput10.VALUE
            If I = 11 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput11.VALUE
            If I = 12 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput12.VALUE
            If I = 13 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput13.VALUE
            If I = 14 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput14.VALUE
            If I = 15 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput15.VALUE
            If I = 16 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput16.VALUE
            If I = 17 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput17.VALUE
            If I = 18 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput18.VALUE
            If I = 19 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput19.VALUE
            If I = 20 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput20.VALUE
            If I = 21 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput21.VALUE
            If I = 22 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput22.VALUE
            If I = 23 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput23.VALUE
            If I = 24 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput24.VALUE
            If I = 25 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput25.VALUE
            If I = 26 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput26.VALUE
            If I = 27 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput27.VALUE
            If I = 28 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput28.VALUE
            If I = 29 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput29.VALUE
            If I = 30 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput30.VALUE
            If I = 31 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput31.VALUE
            If I = 32 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput32.VALUE
            If I = 33 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput33.VALUE
            If I = 34 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput34.VALUE
            If I = 35 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput35.VALUE
            If I = 36 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput36.VALUE
            If I = 37 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput37.VALUE
            If I = 38 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput38.VALUE
            If I = 39 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput39.VALUE
            If I = 40 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput40.VALUE
            If I = 41 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput41.VALUE
            If I = 42 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput42.VALUE
            If I = 43 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput43.VALUE
            If I = 44 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput44.VALUE
            If I = 45 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput45.VALUE
            If I = 46 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput46.VALUE
            If I = 47 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput47.VALUE
            If I = 48 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput48.VALUE
            If I = 49 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput49.VALUE
            If I = 50 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput50.VALUE
            If I = 51 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput51.VALUE
            If I = 52 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput52.VALUE
            If I = 53 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput53.VALUE
            If I = 54 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput54.VALUE
            If I = 55 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput55.VALUE
            If I = 56 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput56.VALUE
            If I = 57 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput57.VALUE
            If I = 58 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput58.VALUE
            If I = 59 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput59.VALUE
            If I = 60 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput60.VALUE
            If I = 61 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput61.VALUE
            If I = 62 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput62.VALUE
            If I = 63 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput63.VALUE
            If I = 64 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput64.VALUE
            If I = 65 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput65.VALUE
            If I = 66 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput66.VALUE
            If I = 67 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput67.VALUE
            If I = 68 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput68.VALUE
            If I = 69 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput69.VALUE
            If I = 70 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput70.VALUE
            If I = 71 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput71.VALUE
            If I = 72 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput72.VALUE
            If I = 73 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput73.VALUE
            If I = 74 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput74.VALUE
            If I = 75 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput75.VALUE
            If I = 76 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput76.VALUE
            If I = 77 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput77.VALUE
            If I = 78 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput78.VALUE
            If I = 79 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput79.VALUE
            If I = 80 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput80.VALUE
            If I = 81 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput81.VALUE
            If I = 82 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput82.VALUE
            If I = 83 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput83.VALUE
            If I = 84 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput84.VALUE
            If I = 85 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput85.VALUE
            If I = 86 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput86.VALUE
            If I = 87 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput87.VALUE
            If I = 88 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput88.VALUE
            If I = 89 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput89.VALUE
            If I = 90 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput90.VALUE
            If I = 91 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput91.VALUE
            If I = 92 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput92.VALUE
            If I = 93 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput93.VALUE
            If I = 94 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput94.VALUE
            If I = 95 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput95.VALUE
            If I = 96 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput96.VALUE
            If I = 97 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput97.VALUE
            If I = 98 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput98.VALUE
            If I = 99 Then checkFeuille = ThisWorkbook.Worksheets(g_CONTROL).cbSelInput99.VALUE
            If checkFeuille Then
                nbFeuilles = nbFeuilles + 1
                setInput (feuille)
            End If
        End If
    Next I
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille de données n'est sélectionnée"
    End If
End Sub
Sub setSheetSel()
    Dim feuille As String
    Dim checkFeuille As Boolean
    Dim nbFeuilles As Integer
    nbFeuilles = 0
    Dim nbMax As Integer
    nbMax = 99
    
    ' Boucler sur les noms dans la feuille
    Set g_WB_Extra = ActiveWorkbook
    ' Détermination de la liste des feuilles à générer
    Dim listSheet() As String
    ReDim listSheet(1 To nbMax)
    Dim ind As Integer
    For I = 1 To nbMax
        ind = I
        checkFeuille = getSheetSelected(ind)
        If getSheetSelected(ind) Then
            listSheet(I) = Trim(ActiveSheet.Cells(g_CONTROL_DATA_L - 1 + I, g_CONTROL_DATA_C).VALUE)
        Else
            listSheet(I) = ""
        End If
    Next I
    Dim FLCIB As Worksheet
    If Not openFileIfNot("TARGET") Then
        MsgBox "Fichier cible inexistant"
        Exit Sub
    End If
    For I = 1 To nbMax
        If listSheet(I) <> "" Then
            Set FLCIB = setOrCreateWsfromName(listSheet(I), g_WB_Target, "CIBLE")
        End If
    Next I
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    g_WB_Extra.Worksheets(g_CONTROL).Select
    For I = 1 To nbMax
        If listSheet(I) <> "" Then
            nbFeuilles = nbFeuilles + 1
            Call DisableExcelSoft
            setSheet (listSheet(I))
            Call EnableExcelSoft
            g_WB_Extra.Worksheets(g_CONTROL).Activate
            g_WB_Extra.Worksheets(g_CONTROL).Select
        End If
    Next I
    Application.StatusBar = "Les feuilles ont été générées"
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille du modèle n'est sélectionnée"
    End If
End Sub
Sub setMetaToSheetSel()
    Call DisableExcel
    Dim sngChrono As Single
    sngChrono = Timer
    Dim FLCTL As Worksheet
    Set FLCTL = Worksheets(g_CONTROL)
    Dim feuille As String
    Dim checkFeuille As Boolean
    Dim nbFeuilles As Integer
    Dim nbErrors As Integer
    nbFeuilles = 0
    Dim nbMax As Integer
    nbMax = 20
    ' Boucler sur les noms dans la feuille
    Set g_WB_Extra = ActiveWorkbook
    ' Détermination de la liste des feuilles à générer
    Dim listSheet() As String
    ReDim listSheet(1 To nbMax)
    Dim ind As Integer
    For I = 1 To nbMax
        ind = I
        checkFeuille = getMetaSheetSelected(ind)
        If getMetaSheetSelected(ind) Then
            listSheet(ind) = Trim(ActiveSheet.Cells(g_CONTROL_META_L - 1 + ind, g_CONTROL_META_C).VALUE)
        Else
            listSheet(ind) = ""
        End If
    Next
    Dim FLCIB As Worksheet
    If Not openFileIfNot("MODELE") Then
        MsgBox "Fichier cible inexistant"
        Exit Sub
    End If
    Dim FLCME As Worksheet
    Set FLCME = setWS(g_CRMETA, g_WB_Modele)
    Call clearLogNom(FLCME)
    Dim newLogLine() As Variant
    newLogLine = Array("META", "DEBUT", "", "DUPLICATION", time(), "", "")
    logNom = alimLog(logNom, newLogLine)
    Dim posFeuille As Integer
    Dim Dnamesgen() As String
    ReDim Dnamesgen(0)
    For ii = 1 To nbMax
        If listSheet(ii) <> "" Then
            nbFeuilles = nbFeuilles + 1
            g_WB_Extra.Worksheets(g_CONTROL).Activate
            g_WB_Extra.Worksheets(g_CONTROL).Select
            posFeuille = ii
            Dnamesgen = appendLists(Dnamesgen, setMetaSheet(listSheet(ii), posFeuille))
            FLCTL.Cells(g_CONTROL_META_CR_L + ii - 1, g_CONTROL_META_CR_C).VALUE = "1"
        End If
    Next
    Call listUpdateGenFromModelSheet(Dnamesgen)
    g_WB_Extra.Worksheets(g_CONTROL).Activate
    g_WB_Extra.Worksheets(g_CONTROL).Select
    sngChrono = Timer - sngChrono
    newLogLine = Array("META", "FIN", "", "DUPLICATION", time(), "", "", "", "Durée = " & "Durée = " & (Int(1000 * sngChrono) / 1000) & " s")
    logNom = alimLog(logNom, newLogLine)
    nbErrors = ecritureLog(FLCME)
    Application.StatusBar = "Les feuilles ont été dupliquées"
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille du modèle n'est sélectionnée"
    End If
    If nbErrors > 0 Then
        MsgBox "Il y a des erreurs dans la duplication" & Chr(10) & "Regardez la feuille de compte-rendu " & g_CRMETA
    End If
    Call EnableExcel
End Sub
Sub setInput(sheet2process As String)
    openFileIfNot ("MODELE")
    ' Alimentation de la feuille QUANTITY avec les positions des quantités en input de la feuille sheet2process
Application.StatusBar = "Traitement de " & nameSheetEnCours
    sngChrono = Timer
    Dim FLICI As Worksheet
    Dim FLCTL As Worksheet
    Dim FLCDG As Worksheet
    Dim FLQUA As Worksheet
    Set FLCTL = ThisWorkbook.Worksheets(g_CONTROL)
    Set FLCDG = g_WB_Modele.Worksheets("CRINPUT")
    Set FLICI = g_WB_Modele.Worksheets(sheet2process)
    Set FLQUA = g_WB_Modele.Worksheets("QUANTITY")
    ' Lecture des quantités
    Dim qua() As Variant
    Dim Quantities() As String
    Dim quantities2() As String
    Dim derlig As Integer
    Dim dercol As Integer
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLQUA.Range("a1:" & DecAlph(dercol) & derlig)
        ReDim qua(1 To derlig, 1 To dercol)
        qua = .VALUE
    End With
    ReDim Quantities(1 To UBound(qua, 1))
    For I = LBound(qua, 1) To UBound(qua, 1)
        Quantities(I) = "[" & qua(I, 3) & "][" & qua(I, 4) & "][" & qua(I, 5) & "]." & qua(I, 6)
    Next
    'Alimentation de Cla tableau de la feuille en cours
    Dim Cla() As Variant
    derlig = FLICI.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLICI.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLICI.Range("a1:" & DecAlph(dercol) & derlig)
        ReDim ClaLig(1 To derlig, 1 To dercol)
        Cla = .VALUE
    End With
    Dim dates() As String
    ReDim dates(1 To dercol)
    Dim ok As Boolean
    Dim aecrire As String
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        ' Détection d'une quantité
        ok = False
        If InStr(Trim(Cla(Nolig, 1)), "[") > 0 Then
            For q = LBound(Quantities) To UBound(Quantities)
                If Quantities(q) = Trim(Cla(Nolig, 1)) Then
                    ok = True
                    aecrire = "@" & sheet2process
                    Exit For
                End If
            Next
        End If
        For NoCol = LBound(Cla, 2) To UBound(Cla, 2)
            ' alimentation du vecteur des dates
            If InStr(Trim(Cla(Nolig, NoCol)), "$") > 0 Then
                dates(NoCol) = Trim(Cla(Nolig, NoCol))
            End If
            If ok Then
                If IsNumeric(Cla(Nolig, NoCol)) And Trim(Cla(Nolig, NoCol)) <> "" Then
                    aecrire = aecrire & ";" & dates(NoCol) & ":" & Nolig & "," & NoCol
                End If
                If NoCol = UBound(Cla, 2) Then
                    qua(Nolig, 17) = qua(Nolig, 17) & aecrire
                End If
            End If
        Next
    Next

    'newLogLine = Array("DATAINP", "FIN", FLICI.NAME, "GENERATION", time(), "", "", "", (Int(1000 * sngChrono) / 1000) & " s")
    'logNom = alimLog(logNom, newLogLine)
    'logNom = Application.Transpose(logNom)
    
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    FLQUA.Range("A1:Q" & derlig).VALUE = qua

    Application.StatusBar = False
End Sub
Sub setSheetComparaison()
    nbFeuilles = 0
    Dim enCours As String
    Dim enCours0 As String
    Dim enCours1 As String
    Dim enCoursd As String
    Dim test1 As Boolean
    Dim testd As Boolean
    Dim precision As Double
    Dim feuille As String
    Dim ind As Integer
    ' Remise à vide des CR
    ' ???
    ' Boucler sur les noms dans la feuille
    For I = 1 To 99
        feuille = Trim(ActiveSheet.Cells(g_CONTROL_DATA_L - 1 + I, g_CONTROL_DATA_C).VALUE)
        If feuille <> "" Then
            ind = I
            checkFeuille = getSheetSelected(ind)
            ' Traiter le cas où fichier modèle <> fichier cible
            'If Left(feuille, 1) = "0" Then feuille = Mid(feuille, 2)
'MsgBox checkFeuille & ":" & i & ":" & feuille
            If checkFeuille Then
                nbFeuilles = nbFeuilles + 1
                'enCours = Mid(feuille, 2)
                'enCours1 = "1" & Mid(feuille, 2)
                'enCoursd = "d" & Mid(feuille, 2)
                precision = ActiveSheet.Cells(g_CONTROL_PRE_L, g_CONTROL_PRE_C).VALUE
                Call setComparaison(feuille, precision)
            End If
        End If
    Next I
    If nbFeuilles = 0 Then
        MsgBox "Aucune feuille du modèle n'est sélectionnée"
    End If
End Sub
Sub setAASheet()
    Set sht = ThisWorkbook.Worksheets("CONTROL")
    Call setSheet("XX")
End Sub
Sub setAAFormula()
    Set sht = ThisWorkbook.Worksheets("CONTROL")
    Call setFormula("XX")
End Sub

Sub setHDSheet()
    Call setSheet("0HD")
End Sub
Sub setHISheet()
    Call setSheet("0HI")
End Sub
Sub setHMSheet()
    Call setSheet("0HM")
End Sub
Sub setHRSheet()
    Call setSheet("0HR")
End Sub
Sub setCRSheet()
    Call setSheet("0CR")
End Sub
Function getPosTerme(term As String, quant() As String, qua() As Variant, times() As String) As String()
    Dim res() As String
    ReDim res(0)
    Dim resSub() As String
    ReDim resSub(0)
    Dim listQuantity() As String
    listQuantity = isAquantity(term, quant)
    Dim iqua As Integer
    Dim timqua As String
    Dim timeMatch As String
    Dim timocc As String
    Dim when As String
    Dim subs As String
    Dim whenList() As String
    Dim pos As String
    Dim she As String
    Dim whespl() As String
    Dim whespl1 As String
    Dim listCoor() As String
    getPosTerme = res
    timeMatch = Split(Split(term & "(", "(")(1) & ")", ")")(0)
    Dim resun As String
    Dim resDate As String
    Dim resDateDroite As String
    Dim newper As String
    Dim newqua As String
    testCode = False
    
'If InStr(term, "[SECTEUR>INDUSTRIE>BRANCHE>Construction.INTENSITE>Diffus.ENERGIE>Electricité:Elec][a][s].conso_spe($h>last") > 0 Then
    'MsgBox UBound(listQuantity) & " getPosTerme" & Chr(10) & term & Chr(10) & Join(listQuantity, Chr(10))
'End If
'If InStr(term, "2015-1") > 0 Then MsgBox term & Chr(10) & Join(listQuantity, Chr(10))
    If UBound(listQuantity) > 0 Then
        iqua = 0
        she = ""
        pos = ""
        For l = LBound(listQuantity) To UBound(listQuantity)
            timqua = qua(CInt(Split(listQuantity(l), "@")(0)), 7)
'If timeMatch = "" Then MsgBox term
            timocc = Split(timeMatch, ">")(0)
            pos = ""
            she = ""
'If InStr(term, "[SECTEUR>INDUSTRIE>BRANCHE>Construction.INTENSITE>Diffus.ENERGIE>Electricité:Elec][a][s].conso_spe($h>last") > 0 Then
'MsgBox listQuantity(l) & Chr(10) & term & Chr(10) & timeMatch & ":" & timqua & Chr(10) & ">>>" & dateInDate(timeMatch, timqua, times) & "<<<"
    'testCode = True
'End If
            resDate = dateInDate(timeMatch, timqua, times)
'MsgBox term & Chr(10) & timeMatch & Chr(10) & timqua & Chr(10) & resDate
'If InStr(term, "2015-1") > 0 Then MsgBox timqua & Chr(10) & timeMatch & Chr(10) & "resDate=" & resDate
'If InStr(term, "[SECTEUR>INDUSTRIE>BRANCHE>Construction.INTENSITE>Diffus.ENERGIE>Electricité:Elec][a][s].conso_spe($h>last") > 0 Then
'If InStr(term, "2015-1") > 0 Then MsgBox subs & Chr(10) & listQuantity(L) & Chr(10) & term & Chr(10) & timeMatch & ":" & timqua & Chr(10) & ">>>" & resDate & "<<<"
    
    'testCode = False
'End If
'MsgBox resDate
            If resDate <> "" Then
                resDateDroite = Split(resDate, ">")(1)
                iqua = CInt(Split(listQuantity(l), "@")(0))
                when = Trim(qua(iqua, 17))
                subs = Trim(qua(iqua, 13))
'If InStr(term, "REPRISE].Eff_e($p>2015-1") > 0 Then
    'MsgBox listQuantity(l) & Chr(10) & when & Chr(10) & subs & ":" & timqua & Chr(10) & ">>>" & dateInDate(timeMatch, timqua, times) & "<<<"
'End If
                If when = "" Then
                    '''Exit For
                    ' pb si la position n'est pas renseignée???
                Else
                    whenList = Split(when, "@")
                    For w = LBound(whenList) To UBound(whenList)
                        If whenList(w) <> "" Then
                            whenlistlist = Split(whenList(w), ";")
                            she = whenlistlist(LBound(whenlistlist))
                            For d = LBound(whenlistlist) To UBound(whenlistlist)
                                If d > LBound(whenlistlist) Then
                                    whespl = Split(Split(whenlistlist(d), ":")(0), ">")
                                    If UBound(whespl) > 0 Then
                                        whespl1 = whespl(1)
                                        If whespl1 = resDateDroite Then
                                            resun = "'" & she & "'!" & Split(whenlistlist(d), ":")(1)
                                            If UBound(res) = 0 Then
                                                ReDim res(1 To 1)
                                            Else
                                                ReDim Preserve res(1 To UBound(res) + 1)
                                            End If
                                            res(UBound(res)) = resun
                                        End If
                                    Else
                                        If resDateDroite = "" Then
                                            resun = "'" & she & "'!" & Split(whenlistlist(d), ":")(1)
                                            If UBound(res) = 0 Then
                                                ReDim res(1 To 1)
                                            Else
                                                ReDim Preserve res(1 To UBound(res) + 1)
                                            End If
                                            res(UBound(res)) = resun
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
                If (UBound(res) = 0 Or when = "") And subs <> "" Then
                    ' on reconstitue les périmètres de la quantité de substitution
'If InStr(term, "REPRISE].Eff_e($p>2015-1") > 0 Then
    'MsgBox listQuantity(l)
'End If
                    splPerS = Split(Split(Split(subs & "(", "(")(0), "].")(0), "]")
                    splPerQ = Split(Split(Split(Split(listQuantity(l) & "(", "(")(0), "].")(0), "@")(1), "]")
                    newper = ""
                    For s = LBound(splPerS) To UBound(splPerS)
                        If splPerS(s) = "[" Then
                            splPerS(s) = splPerQ(s)
                        End If
                        newper = newper & splPerS(s) & "]"
                    Next
                    'On Error GoTo errorHandler
                    newqua = newper & "." & Split(Split(subs & "(", "(")(0), ".")(1) & "(" & timeMatch & ")"
                    ' récursivité
If InStr(testPileRecursivite & ",", "," & (1 + iqua) & ",") > 0 Then
    testPileRecursivite = testPileRecursivite & "," & (1 + iqua)
    If MsgBox("Les quantités semblent boucler sur :" & Chr(10) & Mid(testPileRecursivite, 2) & Chr(10) & term & Chr(10) & "Voulez-vous interrompre les traitements ?", vbYesNo, "") = vbYes Then
        ReDim res(0)
        getPosTerme = res
        End
    End If
End If
testPileRecursivite = testPileRecursivite & "," & (1 + iqua)
                    res = getPosTerme(newqua, quant, qua, times)
'If InStr(term, "REPRISE].Eff_e($p>2015-1") > 0 Then
    'MsgBox listQuantity(l)
'End If
                End If
                Exit For
            Else
            ' Cas ou pas de date et subst à conserver quand la boucle est finie
                iqua = CInt(Split(listQuantity(l), "@")(0))
                If iqua > 0 Then
                    subs = Trim(qua(iqua, 13))
                Else
                    subs = ""
                End If
                If subs <> "" Then
'MsgBox L & Chr(10) & listQuantity(L) & Chr(10) & subs
                    splPerS = Split(Split(Split(subs & "(", "(")(0), "].")(0), "]")
                    splPerQ = Split(Split(Split(Split(listQuantity(l) & "(", "(")(0), "].")(0), "@")(1), "]")
                    newper = ""
                    For s = LBound(splPerS) To UBound(splPerS)
                        If splPerS(s) = "[" Then
                            splPerS(s) = splPerQ(s)
                        End If
                        newper = newper & splPerS(s) & "]"
                    Next
                    newqua = newper & "." & Split(Split(subs & "(", "(")(0), ".")(1) & "(" & timeMatch & ")"
                    ' récursivité
If InStr(testPileRecursivite & ",", "," & (1 + iqua) & ",") > 0 Then
    testPileRecursivite = testPileRecursivite & "," & (1 + iqua)
    If MsgBox("Les quantités semblent boucler sur :" & Chr(10) & Mid(testPileRecursivite, 2) & Chr(10) & term & Chr(10) & Chr(10) & subs & Chr(10) & "Voulez-vous interrompre les traitements ?", vbYesNo, "") = vbYes Then
        ReDim res(0)
        getPosTerme = res
        End
    End If
End If
testPileRecursivite = testPileRecursivite & "," & (1 + iqua)
                    resSub = getPosTerme(newqua, quant, qua, times)
                End If
            End If
        Next
        If UBound(res) = 0 Then
            ' on alimente par le substitute s'il existe
            If UBound(resSub) > 0 Then
'MsgBox Join(resSub, Chr(10))
                res = resSub
            End If
        End If
        'If she <> "" And pos <> "" Then
            'listCoor = Split(pos, ",")
            'getPosTerme = "'" & she & "'!" & DecAlph(CInt(listCoor(UBound(listCoor)))) & listCoor(LBound(listCoor))
        'Else
        getPosTerme = res
        'End If
    Else
        'MsgBox term
    End If
    Exit Function
'errorHandler: MsgBox term & Chr(10) & newper & Chr(10) & subs & Chr(10) & timeMatch
End Function
''' Construction des formules à partir des équations
Sub setFormula(sheet2process As String)
    On Error GoTo errorHandler
    Dim modelName As String
    openFileIfNot ("MODELE")
    openFileIfNot ("TARGET")
    Dim FLICI As Worksheet
    Dim FLCIB As Worksheet
    Dim FLGEN As Worksheet
    Dim FLCTL As Worksheet
    Dim FLCDG As Worksheet
    Dim FLQUA As Worksheet
    Set FLCTL = g_WB_Extra.Worksheets(g_CONTROL)
    Dim nameSheetEnCours As String
    Dim nameTime As String
    Dim nameArea As String
    Dim nameScenario As String
    Dim derlig As Integer
    Dim dercol As Integer
    targetName = getNameIfExists(getNameFromModel(sheet2process, "CIBLE"), g_WB_Target)
    If targetName = "" Then GoTo errorHandlerWsNotExists
    modelName = sheet2process
    If Not WsExist(modelName, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    sngChrono = Timer
    Dim operateur As String
    If g_withoutFormula Then
        operateur = "F@"
    Else
        operateur = "="
    End If
    'If ActiveSheet.NAME = "CONTROL" Then
        nameTime = ActiveSheet.cbTime.VALUE
        nameArea = ActiveSheet.cbArea.VALUE
        nameScenario = ActiveSheet.cbScenario.VALUE
        'If sheet2process = "XX" Then
            'numnom = CInt(FLCTL.Cells(1, 7).VALUE)
            'nameSheetEnCours = Trim(FLCTL.Cells(1 + numnom, 7).VALUE)
            'Set FLCDG = Worksheets("CRFORMULA")
            'Set FLCIB = Worksheets("DATA")
        'Else
            nameSheetEnCours = sheet2process
   '' If Not WsExist(g_CRFORMULA, g_WB_Modele) Then
        ''g_WB_Modele.Worksheets.Add.Move After:=g_WB_Modele.Worksheets(g_WB_Modele.Worksheets.Count)
        ''g_WB_Modele.Worksheets(g_WB_Modele.Worksheets.Count).NAME = g_CRFORMULA
        ''g_WB_Modele.Worksheets(g_CRFORMULA).Range("A1").EntireRow.Insert
        ''For i = LBound(g_listHeadCR) To UBound(g_listHeadCR)
            ''g_WB_Modele.Worksheets(g_CRFORMULA).Cells(1, i).VALUE = g_listHeadCR(i)
        ''Next
        ''g_WB_Extra.Worksheets(g_CONTROL).Activate
        ' ajouter la feuille CONTEXTE TYPE FEUILLE TRAITEMENT TIME ELEMENT OCCURRENCE LIGNE DESCRIPTION
    ''End If
    Set FLCDG = setWS(g_CRFORMULA, g_WB_Modele)
    If FLCDG.Cells(1, 1).VALUE = "" Then
        For I = LBound(g_listHeadCR) To UBound(g_listHeadCR)
            g_WB_Modele.Worksheets(g_CRFORMULA).Cells(1, I).VALUE = g_listHeadCR(I)
        Next
    End If
            ''Set FLCDG = g_WB_Modele.Worksheets(g_CRFORMULA)
    'Set FLCIB = setWS(Mid(nameSheetEnCours, 2), g_WB_Target)
    Set FLCIB = g_WB_Target.Worksheets(targetName)
        'End If
    'Else
        'nameSheetEnCours = ActiveSheet.NAME
        'Set FLCDG = Worksheets("CR" & nameSheetEnCours)
        'Set FLCIB = Worksheets(Mid(nameSheetEnCours, 2))
    'End If
    'Set FLICI = setWS(nameSheetEnCours, g_WB_Modele)
    Set FLICI = g_WB_Modele.Worksheets(modelName)
Application.StatusBar = "Traitement de " & nameSheetEnCours
    'derlig = Split(FLCDG.UsedRange.Address, "$")(4)
    derlig = getDerLig(FLCDG)
    For Nolig = 1 To derlig
        If FLCDG.Cells(Nolig, 1) <> "" Then LigDebCNo = Nolig + 1
    Next
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = FLCDG.Range("A" & FLCDG.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCDG.Columns.Count).End(xlToLeft).Column
    With FLCDG.Range("a1:" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Dim newLogLine() As Variant
newLogLine = Array("DATAFORMUL", "DEBUT", FLICI.NAME, "GENERATION", time())
logNom = alimLog(logNom, newLogLine)
    'If Not WsExist(nameSheetEnCours, g_WB_Modele) Then
'newLogLine = Array("", "ERREUR", FLICI.NAME, "IMPORTATION", time(), "", "", "", "FEUILLE GENERIQUE INEXISTANTE")
'logNom = alimLog(logNom, newLogLine)
        'MsgBox "La feuille " & nameSheetEnCours & " n'exite pas dans le modèle"
        'GoTo errorHandler2
    'Else
        Set FLGEN = g_WB_Modele.Worksheets(modelName)
    'End If
    If Not WsExist("QUANTITY", g_WB_Modele) Then
        GoTo errorHandler3
    Else
        Set FLQUA = g_WB_Modele.Worksheets("QUANTITY")
    End If
    ' Lecture des quantités
    Dim qua() As Variant
    Dim Quantities() As String
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLQUA.Range("a" & 2 & ":" & DecAlph(dercol) & derlig)
        ReDim qua(1 To (derlig - 1), 1 To dercol)
        qua = .VALUE
    End With
    ' lecture de TIME
    Dim times() As String
    times = lectureTime(nameTime)
    ' Désactivation des formules
    ''''''Application.Calculation = xlCalculationManual
    ''''''Application.ScreenUpdating = False
    '''Dim PosFct As Integer
    '''Dim LigFct As Integer
    '''NOMBRE = 0
    '''NbEquations = 0
    '''''derlig = FLCIB.Cells.SpecialCells(xlCellTypeLastCell).Row
    '''''dercol = FLCIB.Cells(1, Columns.Count).End(xlToLeft).Column
    'derlig = FLCIB.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    derlig = getDerLig(FLCIB)
    'dercol = FLCIB.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    dercol = getDerCol(FLCIB)
    '''LigCal = GetFirstLine(FLCAL)
    '''ColCal = GetFirstCol(FLCAL)
    '''PosFct = ColCal(1)
    Dim fomulaOf() As Variant
    Dim fomulaOfOld() As Variant
    fomulaOf = FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal
    fomulaOfOld = FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal
    Dim ListTermes() As String
    Dim formul As String
    Dim formula As String
    Dim posTerme() As String
    Dim getPosTermeOk As String
    Dim term As String
    Dim numPosTerme As Integer
    Dim formulencours As String
    ReDim Quantities(1 To UBound(qua, 1))
    For I = LBound(qua, 1) To UBound(qua, 1)
        Quantities(I) = "[" & qua(I, 3) & "][" & qua(I, 4) & "][" & qua(I, 5) & "]." & qua(I, 6)
    Next
    compteur = 0
    lastCompteur = 0
    Dim posPost As Integer
    Dim valdef As Integer
    Dim enPLus1 As String
    enPLus1 = ""
    Dim enPLus2 As String
    enPLus2 = ""
    Dim enPLusBool As Boolean
    enPLusBool = True
'MsgBox LBound(fomulaOf, 1) & ":" & UBound(fomulaOf, 1) & Chr(10) & LBound(fomulaOf, 2) & ":" & UBound(fomulaOf, 2)
    For Nolig = LBound(fomulaOf, 1) To UBound(fomulaOf, 1)
        For NoCol = LBound(fomulaOf, 2) To UBound(fomulaOf, 2)
compteur = Int(100 * ((Nolig - 1) * UBound(fomulaOf, 2) + NoCol) / (UBound(fomulaOf, 1) * UBound(fomulaOf, 2)))
If compteur <> lastCompteur Then
Application.StatusBar = "FORMULES DE " & nameSheetEnCours & " : GENERATION : " & ((Nolig - 1) * UBound(fomulaOf, 2) + NoCol) & " / " & (UBound(fomulaOf, 1) * UBound(fomulaOf, 2)) & " " & compteur & " %"
DoEvents
lastCompteur = compteur
End If
            If Left(fomulaOf(Nolig, NoCol), 2) = "E@" Then
                formul = Split(fomulaOf(Nolig, NoCol), "E@")(1)
                enPLusBool = True
                enPLus = ""
                If InStr(LCase(formul), "interpolation") > 0 Then
                    enPLus1 = ""
                    enPLus2 = ""
                    For p = NoCol To UBound(fomulaOf, 2)
                        posPost = UBound(fomulaOf, 2)
                        If InStr(LCase(fomulaOf(Nolig, p)), "interpolation") = 0 Then
                            posPost = p
                            Exit For
                        End If
                    Next
                    '''" (DecAlph(NoCol - 1) & Nolig) &
'formul = (DecAlph(NoCol - 1) & Nolig) & " + (" & (DecAlph(posPost) & Nolig) & " - " & (DecAlph(NoCol - 1) & Nolig) & ") / " & (posPost - NoCol + 1)
                    enPLus1 = " + (" & (DecAlph(posPost) & Nolig) & " - "
                    enPLus2 = ") / " & (posPost - NoCol + 1)
                    formul = Split(Split(formul, "interpolation")(0) & " + ", " + ")(0)
                    formul = formul & " + 1111111111111111 + " & formul & " + 2222222222222222"
                    If (posPost - NoCol + 1) > 0 Then
                        '''fomulaOf(Nolig, NoCol) = "=" & formul
                    Else
                        enPLusBool = False
                        fomulaOf(Nolig, NoCol) = Replace(Replace(formul, " + 1111111111111111 + ", enPLus1), " + 2222222222222222", enPLus2)
                    End If
                End If
                If enPLusBool Then
'MsgBox "formule=" & formul & enPLus
                ListTermes = analyseTerme(formul)
                formulencours = formul
                If UBound(ListTermes) > 0 Then
'Application.StatusBar = Nolig & ":" & NoCol & ":ListTermes=" & UBound(ListTermes)
'MsgBox "formule=" & formul & enPLus & Chr(10) & Join(ListTermes, Chr(10))
                    ' recherche de la position des termes
                    For t = LBound(ListTermes) To UBound(ListTermes)
                        testPileRecursivite = ""
                        posTerme = getPosTerme(ListTermes(t), Quantities, qua, times)
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme)
                        formulencours = Replace(formulencours, ListTermes(t), "@" & t, 1, 1)
                        If UBound(posTerme) = 0 Then
                            ' on remplace ce qu on a pas trouvé par 1 ou 0 et ALERTE ?
                            ' Détection de l'additivité ou non
                            spla = Split(formulencours, "@" & t)(0)
                            valdef = 1
                            If InStr(LCase(spla), "somme(") > 0 Then
                                spls = Split(LCase(spla), "somme(")(1)
                                If InStr(spls, ")") = 0 Then valdef = 0
                            End If
                            formul = Replace(formul, ListTermes(t), valdef)
                        Else
                            ' prendre le plus proche sauf lui même
'MsgBox formul & Chr(10) & Join(posTerme, Chr(10))
                            numPosTerme = 0
'If fomulaOf(Nolig, NoCol) = "E@[VEHICULES>Véhicules hybrides rechargeables:Gazole,Elec,hyb.SECTEUR>TRANSPORT>MARCHANDISES>MODES>Véhicules utilitaires][a][s].Parc($h>2000)" Then _
'MsgBox Join(ListTermes, Chr(10)) & Chr(10) & Join(posTerme, Chr(10))
                            ' analyse de lui-même
                            For c = LBound(posTerme) To UBound(posTerme)
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme) & ":c=" & c
                                term = posTerme(c)
                                If Trim(term) <> "" Then
                                    termeSpl = Split(term, "!")
                                    If termeSpl(0) = "'" & FLCIB.NAME & "'" Then
                                        termeChoisiSpl = Split(term, "!")
                                        listCoor = Split(termeChoisiSpl(1), ",")
                                        If Nolig & "," & NoCol <> termeChoisiSpl(1) Then
                                            If UBound(ListTermes) > 1 Then
                                            'If False Then
                                                numPosTerme = c
                                                Exit For
                                            Else

                                                ' MAIS POURQUOI ?????
                                                ' cas où une qua est répétée (pour éviter les références circulaires)
                                                ' ne va chercher que la qua avant
                                                ici = CInt(Nolig)
                                                etu = CInt(Split(termeChoisiSpl(1), ",")(0))
'MsgBox formul & Chr(10) & term & Chr(10) & ListTermes(UBound(ListTermes)) & Chr(10) & ici & "::" & etu

                                                If etu < ici Then
                                                    numPosTerme = c
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ' analyse des autres
                            If numPosTerme = 0 Then
                            For c = LBound(posTerme) To UBound(posTerme)
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme) & ":autres c=" & c
                                term = posTerme(c)
                                If Trim(term) <> "" Then
                                    termeSpl = Split(term, "!")
                                    If termeSpl(0) <> "'" & FLCIB.NAME & "'" Then
                                        numPosTerme = c
                                        Exit For
                                    End If
                                End If
                            Next
                            End If
'If fomulaOf(Nolig, NoCol) = "E@[VEHICULES>Véhicules hybrides rechargeables:Gazole,Elec,hyb.SECTEUR>TRANSPORT>MARCHANDISES>MODES>Véhicules utilitaires][a][s].Parc($h>2000)" Then _
'MsgBox Join(ListTermes, Chr(10)) & Chr(10) & Join(posTerme, Chr(10)) & Chr(10) & ">>>" & posTerme(numPosTerme)
                            If numPosTerme > 0 Then
                                termeChoisi = posTerme(numPosTerme)
'MsgBox termeChoisi
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme) & ":termeChoisi=" & termeChoisi
                                termeChoisiSpl = Split(termeChoisi, "!")
                                listCoor = Split(termeChoisiSpl(1), ",")
                                getPosTermeOk = termeChoisiSpl(0) & "!" & DecAlph(CInt(listCoor(UBound(listCoor)))) & listCoor(LBound(listCoor))
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme) & ":getPosTermeOk=" & getPosTermeOk
                                formul = Replace(formul, ListTermes(t), getPosTermeOk)
'Application.StatusBar = Nolig & ":" & NoCol & ":posTerme=" & UBound(posTerme) & ":getPosTermeOk FIN=" & getPosTermeOk
                            Else
                                formul = Replace(formul, ListTermes(t), 1)
                            End If
                        End If
                    Next
'Application.StatusBar = Nolig & ":" & NoCol
                    formul = Replace(Replace(formul, " + 1111111111111111 + ", enPLus1), " + 2222222222222222", enPLus2)
'MsgBox formul
                    fomulaOf(Nolig, NoCol) = operateur & remplaceDateByVal(formul, times)
'MsgBox Nolig & ":" & NoCol & ":" & formul
'Application.StatusBar = Nolig & ":" & NoCol
                Else
                    'traiter le cas formule = numérique ???
                    formula = remplaceDateByVal(formul, times)
                    'L'idéal serait de tester la validité d'une formule avant son écriture dans une cellule
                    ' ==> écrire dans une cellule ficive et utiliser application.validate ??
                    '''If testFormulaValide(formula) Then
                        '''fomulaOf(Nolig, NoCol) = "=" & formula
                    '''Else
                    fomulaOf(Nolig, NoCol) = formula
                    '''End If
                End If
                End If
            End If
            If Left(fomulaOf(Nolig, NoCol), 2) = "F@" And Not g_withoutFormula Then
                fomulaOf(Nolig, NoCol) = operateur & Split(fomulaOf(Nolig, NoCol), "F@")(1)
            End If
'Application.StatusBar = Nolig & ":" & NoCol & "::" & "end"
        Next
    Next
'MsgBox "0"
    'For NoCol = LBound(fomulaOf, 2) To UBound(fomulaOf, 2)
        'MsgBox Trim(fomulaOf(4, NoCol))
    'Next
''MsgBox fomulaOf
    '''On Error GoTo errorHandler1
'MsgBox LBound(fomulaOf, 1) & ":" & UBound(fomulaOf, 1) & Chr(10) & LBound(fomulaOf, 2) & ":" & UBound(fomulaOf, 2)
'MsgBox LBound(FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal, 1) & ":" & UBound(FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal, 1) & Chr(10) & LBound(FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal, 2) & ":" & UBound(FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal, 2)
    '''Dim iiiii As Integer
    '''Dim jjjjj As Integer
'''On Error GoTo errorHandlerFormula
    '''For iiiii = LBound(fomulaOf, 1) To UBound(fomulaOf, 1)
        '''For jjjjj = LBound(fomulaOf, 2) To UBound(fomulaOf, 2)
'MsgBox iiiii & ":" & jjjjj

            '''FLCIB.Range(DecAlph(jjjjj) & iiiii).FormulaLocal = fomulaOf(iiiii, jjjjj)
        '''Next
    '''Next
'MsgBox "puis"
    FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal = fomulaOf
    ''''''Application.Calculation = xlCalculationAutomatic
    ''''''Application.ScreenUpdating = True
    sngChrono = Timer - sngChrono
    'newLogLine = Array("", "INFO", FLICI.NAME, "Statistiques", Time(), "", "", nbliginit & " lignes génériques", nbliggener & " lignes générées")
'logNom = alimLog(logNom, newLogLine)
    'newLogLine = Array("", "INFO", FLICI.NAME, "Statistiques", Time(), "", "", nbcolinit & " colonnes génériques", nbcolgener & " colonnes générées")
'logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("DATAFORMUL", "FIN", FLICI.NAME, "GENERATION", time(), "", "", "", (Int(1000 * sngChrono) / 1000) & " s")
logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCDG.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    FLCDG.Range("A1:K" & UBound(logNom, 1)).CurrentRegion.Borders.LineStyle = xlContinuous
    resColoriage = 0

    'If FLCTL.Cells(9, 15).VALUE = "" Then
        'FLCTL.Cells(9, 15).VALUE = "" & nbliginit & ";" & nbcolinit
        'FLCTL.Cells(9, 16).VALUE = "" & nbliggener & ";" & nbcolgener
        'FLCTL.Cells(9, 17).VALUE = "" & resColoriage
        'FLCTL.Cells(9, 18).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
    'Else
        'FLCTL.Cells(9, 15).VALUE = FLCTL.Cells(9, 15).VALUE & Chr(10) & ""
        'FLCTL.Cells(9, 16).VALUE = FLCTL.Cells(9, 16).VALUE & Chr(10) & ""
        'FLCTL.Cells(9, 17).VALUE = FLCTL.Cells(9, 17).VALUE & Chr(10) & resColoriage
        'FLCTL.Cells(9, 18).VALUE = FLCTL.Cells(9, 18).VALUE & Chr(10) & Round((Int(1000 * sngChrono) / 1000))
    'End If
    l = getLineFrom(sheet2process, FLCTL, g_CONTROL_DATA_L, g_CONTROL_DATA_C)
    FLCTL.Cells(l, g_CONTROL_DATA_FOR_C).VALUE = "" & nbcolinit
    FLCTL.Cells(l, g_CONTROL_DATA_FOR_C + 1).VALUE = "" & nbcolgener
    FLCTL.Cells(l, g_CONTROL_DATA_FOR_C + 2).VALUE = "" & resColoriage
    FLCTL.Cells(l, g_CONTROL_DATA_FOR_C + 3).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
Application.StatusBar = False
Exit Sub
errorHandler3: Call onErrDo("La feuille " & "QUANTITY" & " n'exite pas dans le modèle", "setFormula"): Exit Sub
errorHandler2: Call onErrDo("La feuille " & nameSheetEnCours & " n'exite pas dans le modèle", "setFormula"): Exit Sub
errorHandler1: Call onErrDo("Attention certaines formules sont invalides", "setFormula"): Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setFormula"): Exit Sub
errorHandlerWsNotExists: Call onErrDo("La feuille n'existe pas", "setFormula"): Exit Sub
'''errorHandlerFormula: Call onErrDo(fomulaOf(iiiii, jjjjj) & Chr(10) & "Il y a des erreurs : L=", iiiii & ":C=" & jjjjj & " n°"): Exit Sub
    'indique le numéro et la description de l'erreur survenue
    'divzero = "9999999"
    
    'For Nolig = LBound(fomulaOf, 1) To UBound(fomulaOf, 1)
        'For NoCol = LBound(fomulaOf, 2) To UBound(fomulaOf, 2)
            'fomulaOf(Nolig, NoCol) = "=si(esterr(fomulaOf(Nolig, NoCol));" & divzero & ";fomulaOf(Nolig, NoCol))"
        'Next
    'Next
    '''FLCIB.Range("A1:" & DecAlph(dercol) & derlig).FormulaLocal = fomulaOf
    'Application.Calculation = xlCalculationAutomatic
End Sub
Function testFormulaValide(formul As String) As Boolean
    testFormulaValide = Not IsError(Application.Evaluate("=" & formul))
    MsgBox formul & Chr(10) & testFormulaValide
End Function


Sub listDataSuite()
    Dim listFeuilles() As String
    ReDim listFeuilles(0)
    Dim listFeuillesAt() As String
    ReDim listFeuillesAt(0)
    'IsError(Sheets("Feuil1")))
    For I = 0 To ActiveSheet.ListData.ListCount - 1
        If UBound(listFeuilles) = 0 Then
            ReDim listFeuilles(1 To 1)
        Else
            ReDim Preserve listFeuilles(1 To UBound(listFeuilles) + 1)
        End If
        If UBound(listFeuillesAt) = 0 Then
            ReDim listFeuillesAt(1 To 1)
        Else
            ReDim Preserve listFeuillesAt(1 To UBound(listFeuillesAt) + 1)
        End If
        If ActiveSheet.ListData.Selected(I) Then
            If Not WsExist(Mid(ActiveSheet.ListData.list(I), 2)) Then
                listFeuilles(UBound(listFeuilles)) = "?"
            Else
                listFeuilles(UBound(listFeuilles)) = Mid(ActiveSheet.ListData.list(I), 2)
            End If
            If Not WsExist("1" & Mid(ActiveSheet.ListData.list(I), 2)) Then
                listFeuillesAt(UBound(listFeuillesAt)) = "?"
            Else
                listFeuillesAt(UBound(listFeuillesAt)) = "1" & Mid(ActiveSheet.ListData.list(I), 2)
            End If
        Else
            listFeuilles(UBound(listFeuilles)) = ""
            listFeuillesAt(UBound(listFeuillesAt)) = ""
        End If
    Next I
    ActiveSheet.Cells(9, 2).VALUE = Join(listFeuilles, Chr(10))
    ActiveSheet.Cells(9, 3).VALUE = Join(listFeuillesAt, Chr(10))
End Sub
Sub setComparaison(sheet2process As String, prec As Double)
    On Error GoTo errorHandler
    Dim diffName As String
    openFileIfNot ("MODELE")
    openFileIfNot ("SOURCE")
    openFileIfNot ("TARGET")
    sngChrono = Timer
    Dim FLSHEET As Worksheet
    Dim FLSHEET1 As Worksheet
    Dim FLSHEETd As Worksheet
    Dim FLCTL As Worksheet
    targetName = getNameIfExists(getNameFromModel(sheet2process, "CIBLE"), g_WB_Target)
    If targetName = "" Then GoTo errorHandlerWsNotExists
    sourceNAme = getNameIfExists(getNameFromModel(sheet2process, "SOURCE"), g_WB_Source)
    If sourceNAme = "" Then Exit Sub
    diffName = "d" & targetName
    Set FLCTL = Worksheets(g_CONTROL)
    'Set FLSHEET = setWS(targetName, g_WB_Target)
    Set FLSHEET = g_WB_Target.Worksheets(targetName)
    'Set FLSHEET1 = setWS(sourceNAme, g_WB_Source)
    Set FLSHEET1 = g_WB_Source.Worksheets(sourceNAme)
    Set FLSHEETd = setWS(diffName, g_WB_Target)
    'Set FLSHEETd = Worksheets(sheet2processd)
    Dim DAT() As Variant
    Dim DAT1() As Variant
    Dim DATd() As Variant
    Dim derlig As Integer
    Dim dercol As Integer
    derlig = getDerLig(FLSHEET)
    dercol = getDerCol(FLSHEET)
    'derlig = FLSHEET.Cells.SpecialCells(xlCellTypeLastCell).Row
    'dercol = Columns(Split(FLSHEET.UsedRange.Address, "$")(3)).Column
    With FLSHEET.Range("a1:" & DecAlph(dercol) & derlig)
        DAT = .VALUE
    End With
    derlig = getDerLig(FLSHEET1)
    dercol = getDerCol(FLSHEET1)
    'derlig = FLSHEET1.Cells.SpecialCells(xlCellTypeLastCell).Row
    'dercol = Columns(Split(FLSHEET1.UsedRange.Address, "$")(3)).Column
    Dim SuiOld() As Variant
    With FLSHEET1.Range("a1:" & DecAlph(dercol) & derlig)
        SuiOld = .VALUE
        DAT1 = setDataTransport(setDataTertiaire(SuiOld))
        DATd = setDataTransport(setDataTertiaire(SuiOld))
    End With
    Dim nbErrorLig As Integer
    FLSHEETd.Cells.Clear
    For Nolig = LBound(DAT1, 1) To UBound(DAT1, 1)
        nbErrorLig = 0
        For NoCol = LBound(DAT1, 2) To UBound(DAT1, 2)
            If Not IsEmpty(DAT1(Nolig, NoCol)) Then
                If IsNumeric(DAT1(Nolig, NoCol)) Then
                    If Nolig <= UBound(DAT, 1) And NoCol <= UBound(DAT, 2) Then
                        If IsNumeric(DAT(Nolig, NoCol)) Then
                            If DAT1(Nolig, NoCol) <> DAT(Nolig, NoCol) Then
                                If Abs(DAT1(Nolig, NoCol) - DAT(Nolig, NoCol)) > prec Then
                                    DATd(Nolig, NoCol) = "E:" & DAT1(Nolig, NoCol) & ":" & DAT(Nolig, NoCol)
                                    nbErrorLig = nbErrorLig + 1
                                    FLSHEETd.Cells(Nolig, NoCol).Interior.Color = RGB(255, 160, 160)
                                Else
                                    DATd(Nolig, NoCol) = DAT1(Nolig, NoCol)
                                    FLSHEETd.Cells(Nolig, NoCol).Interior.Color = FLSHEET1.Cells(Nolig, NoCol).Interior.Color
                                End If
                            Else
                                DATd(Nolig, NoCol) = DAT1(Nolig, NoCol)
                                FLSHEETd.Cells(Nolig, NoCol).Interior.Color = FLSHEET1.Cells(Nolig, NoCol).Interior.Color
                            End If
                        End If
                    Else
                        DATd(Nolig, NoCol) = ""
                    End If
                End If
            End If
            If NoCol = UBound(DAT1, 2) Then
                If nbErrorLig > 0 Then
                    DATd(Nolig, LBound(DAT1, 2)) = "[" & nbErrorLig & "]" & DATd(Nolig, LBound(DAT1, 2))
                    FLSHEETd.Cells(Nolig, LBound(DAT1, 2)).Interior.Color = RGB(255, 160, 160)
                Else
                    FLSHEETd.Cells(Nolig, LBound(DAT1, 2)).Interior.Color = FLSHEET1.Cells(Nolig, LBound(DAT1, 2)).Interior.Color
                End If
            End If
        Next
    Next
    FLSHEETd.Range("a1:" & DecAlph(dercol) & derlig).VALUE = DATd
    l = getLineFrom(sheet2process, FLCTL, g_CONTROL_DATA_L, g_CONTROL_DATA_C)
    sngChrono = Timer - sngChrono
'MsgBox sheet2process & Chr(10) & FLCTL.NAME & Chr(10) & Round((Int(1000 * sngChrono) / 1000)) & Chr(10) & l & ":" & g_CONTROL_DATA_COM_C
    FLCTL.Cells(l, g_CONTROL_DATA_COM_C + 3).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
    FLCTL.Select
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setComparaison"): Exit Sub
errorHandlerWsNotExists: Call onErrDo("La feuille n'existe pas", "setFormula"): Exit Sub
End Sub
''' Retourne le WS en fonction du nom de la feuille modèle et le crée si nécessaire
Function setOrCreateWsfromName(sheet2process As String, workBookFormModel As Workbook, typ As String) As Worksheet
    Dim namesFromModel As String
    'Dim workBookFormModel As Workbook
    Dim wsEnCours As String
    'If typ = "CIBLE" Then Set workBookFormModel = g_WB_Target
'MsgBox g_WB_Target.NAME
    'If typ = "SOURCE" Then Set workBookFormModel = g_WB_Source
    'If typ = "MODELE" Then Set workBookFormModel = g_WB_Modele
    namesFromModel = getNameFromModel(sheet2process, typ)
    splitNamesFromModel = Split(namesFromModel, "@")
    For I = LBound(splitNamesFromModel) To UBound(splitNamesFromModel)
        wsEnCours = splitNamesFromModel(I)
        If WsExist(wsEnCours, workBookFormModel) Then
            Set setOrCreateWsfromName = workBookFormModel.Worksheets(wsEnCours)
            Exit Function
        End If
    Next
    ' cas où le ws n'existe ps => on le crée
    wsEnCours = splitNamesFromModel(LBound(splitNamesFromModel))
    workBookFormModel.Worksheets.Add.Move After:=workBookFormModel.Worksheets(workBookFormModel.Worksheets.Count)
    workBookFormModel.Worksheets(workBookFormModel.Worksheets.Count).NAME = wsEnCours
    '''g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set setOrCreateWsfromName = workBookFormModel.Worksheets(wsEnCours)
End Function
''' Retourne le WS en fonction du nom de la feuille modèle et le crée si nécessaire
Function setOrCreateWsfromModel(sheet2process As String, typ As String) As Worksheet
    Dim namesFromModel As String
    Dim workBookFormModel As Workbook
    Dim wsEnCours As String
    If typ = "CIBLE" Then Set workBookFormModel = g_WB_Target
'MsgBox g_WB_Target.NAME
    If typ = "SOURCE" Then Set workBookFormModel = g_WB_Source
    If typ = "MODELE" Then Set workBookFormModel = g_WB_Modele
    namesFromModel = getNameFromModel(sheet2process, typ)
    splitNamesFromModel = Split(namesFromModel, "@")
    For I = LBound(splitNamesFromModel) To UBound(splitNamesFromModel)
        wsEnCours = splitNamesFromModel(I)
        If WsExist(wsEnCours, workBookFormModel) Then
            Set setOrCreateWsfromModel = workBookFormModel.Worksheets(wsEnCours)
            Exit Function
        End If
    Next
    ' cas où le ws n'existe ps => on le crée
    wsEnCours = splitNamesFromModel(LBound(splitNamesFromModel))
    workBookFormModel.Worksheets.Add.Move After:=workBookFormModel.Worksheets(workBookFormModel.Worksheets.Count)
    workBookFormModel.Worksheets(workBookFormModel.Worksheets.Count).NAME = wsEnCours
    '''g_WB_Extra.Worksheets(g_CONTROL).Activate
    Set setOrCreateWsfromModel = workBookFormModel.Worksheets(wsEnCours)
End Function

''' Retourne le nom de la feuille si elle existe "" sinon
Function getNameIfExists(names As String, wb As Workbook) As String
    Dim str As String
    splNames = Split(names & "@", "@")
    For I = LBound(splNames) To UBound(splNames)
        If splNames(I) <> "" Then
            str = splNames(I)
            If WsExist(str, wb) Then
                getNameIfExists = splNames(I)
                Exit Function
            End If
        End If
    Next
    getNameIfExists = ""
End Function

''' retourne le nom de la feuille en fonction du nom de la feuille du modèle
Function getNameFromModel(sheet2process As String, typ As String) As String
    ' cas ou tout est dans un seul fichier
    If g_WB_Modele_Name = g_WB_Target_Name And g_WB_Modele_Name = g_WB_Source_Name Then
        If typ = "CIBLE" Then
            getNameFromModel = Mid(sheet2process, 2)
            Exit Function
        End If
        If typ = "SOURCE" Then
            getNameFromModel = "1" & Mid(sheet2process, 2)
            Exit Function
        End If
    End If
    If g_WB_Modele_Name = g_WB_Target_Name And g_WB_Target_Name <> g_WB_Source_Name Then
        If typ = "CIBLE" Then
            getNameFromModel = Mid(sheet2process, 2)
            Exit Function
        End If
        If typ = "SOURCE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = Mid(sheet2process, 2) & "@" & "1" & Mid(sheet2process, 2)
            Else
                getNameFromModel = sheet2process & "@" & "1" & sheet2process
            End If
            Exit Function
        End If
    End If
    If g_WB_Modele_Name = g_WB_Source_Name And g_WB_Target_Name <> g_WB_Source_Name Then
        If typ = "CIBLE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = Mid(sheet2process, 2)
            Else
                getNameFromModel = sheet2process
            End If
            Exit Function
        End If
        If typ = "SOURCE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = "1" & Mid(sheet2process, 2)
            Else
                getNameFromModel = "1" & sheet2process
            End If
            Exit Function
        End If
    End If
    If g_WB_Modele_Name <> g_WB_Target_Name And g_WB_Target_Name = g_WB_Source_Name Then
        If typ = "CIBLE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = Mid(sheet2process, 2)
            Else
                getNameFromModel = sheet2process
            End If
            Exit Function
        End If
        If typ = "SOURCE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = "1" & Mid(sheet2process, 2)
            Else
                getNameFromModel = "1" & sheet2process
            End If
            Exit Function
        End If
    End If
    If g_WB_Modele_Name <> g_WB_Target_Name And g_WB_Modele_Name <> g_WB_Source_Name And g_WB_Target_Name <> g_WB_Source_Name Then
        If typ = "CIBLE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = Mid(sheet2process, 2)
            Else
                getNameFromModel = sheet2process
            End If
            Exit Function
        End If
        If typ = "SOURCE" Then
            If Left(sheet2process, 1) = "0" Then
                getNameFromModel = Mid(sheet2process, 2) & "@" & "1" & Mid(sheet2process, 2)
            Else
                getNameFromModel = sheet2process & "@" & "1" & sheet2process
            End If
            Exit Function
        End If
    End If
End Function
''' Génération des feuilles génériques à partir d'une méta feuille
Function setMetaSheet(sheet2process As String, posFeuille As Integer) As String()
    On Error GoTo errorHandler
    Dim Dnamesgen() As String
    ReDim Dnamesgen(0)
    setMetaSheet = Dnamesgen
    Dim sngChrono As Single
    sngChrono = Timer
    Dim nbErrors As Integer
    nbErrors = 0
    Dim newLogLine() As Variant
    Dim targetName As String
    Dim sourceNAme As String
    openFileIfNot ("MODELE")
    Dim FLICI As Worksheet
    Dim FLCTL As Worksheet
    Set FLCTL = Worksheets(g_CONTROL)
    Dim FLCDG As Worksheet
    Dim FLQUA As Worksheet
    Dim FLAEC As Worksheet
    Dim FLCME As Worksheet
    Set FLCME = setWS(g_CRMETA, g_WB_Modele)
    Dim nameSheetEnCours As String
    Dim nameTime As String
    Dim nameArea As String
    Dim nameScenario As String
    nameTime = ActiveSheet.cbTime.VALUE
    nameArea = ActiveSheet.cbArea.VALUE
    nameScenario = ActiveSheet.cbScenario.VALUE
    nameSheetEnCours = sheet2process
    Set FLCDG = setWS(g_CRDATA, g_WB_Modele)
    If Not WsExist(nameSheetEnCours, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLICI = setWS(nameSheetEnCours, g_WB_Modele)
    If Not WsExist(g_QUANTITY, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLQUA = g_WB_Modele.Worksheets(g_QUANTITY)
    If Not WsExist(g_NOMENCLATURE, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLNOM = g_WB_Modele.Worksheets(g_NOMENCLATURE)
    If Not WsExist(g_AREA, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLAREA = g_WB_Modele.Worksheets(g_AREA)
    If Not WsExist(g_SCENARIO, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLSCENARIO = g_WB_Modele.Worksheets(g_SCENARIO)
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = getDerLig(FLICI)
    DerColSheet = getDerCol(FLICI)
Application.StatusBar = "Traitement de " & nameSheetEnCours
    With FLICI.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
        ReDim Cla(1 To derligsheet, 1 To DerColSheet)
        Cla = .VALUE
    End With
    ' lecture de la nomenclature
    Dim NOMENC() As Variant
    Dim NOMENCLA() As String
    Dim IDNOMENC() As Variant
    Dim IDNOMENCLA() As String
    Dim eten() As String
    ReDim eten(0)
    derlignom = FLNOM.Range("C" & FLNOM.Rows.Count).End(xlUp).Row
    NOMENC = FLNOM.Range("c2:c" & derlignom).VALUE
    ReDim NOMENCLA(1 To UBound(NOMENC, 1))
    For nol = LBound(NOMENC, 1) To UBound(NOMENC, 1)
        NOMENCLA(nol) = NOMENC(nol, LBound(NOMENC, 2))
    Next
    ' lecture des areas
    Dim AREAS() As Variant
    Dim areaList() As String
    derlignom = FLAREA.Range("A" & FLAREA.Rows.Count).End(xlUp).Row
    AREAS = FLAREA.Range("A2:B" & derlignom).VALUE
    ReDim areaList(1 To UBound(AREAS, 1))
    For a = LBound(AREAS, 1) To UBound(AREAS, 1)
        areaList(a) = AREAS(a, 1)
    Next
    ' lecture des scénarios
    Dim SCENARIOS() As Variant
    Dim scenarioList() As String
    derlignom = FLSCENARIO.Range("A" & FLSCENARIO.Rows.Count).End(xlUp).Row
    SCENARIOS = FLSCENARIO.Range("A2:B" & derlignom).VALUE
    ReDim scenarioList(1 To UBound(SCENARIOS, 1))
    For a = LBound(SCENARIOS, 1) To UBound(SCENARIOS, 1)
        scenarioList(a) = SCENARIOS(a, 1)
    Next
    racineSCENARIO = SCENARIOS(1, 1)
    ' Analyse de la 1ère cellule
    Dim listBoucle() As Variant
    ReDim listBoucle(1 To 5)
    Dim listModalites() As String
    ReDim listModalites(0 To 0)
    For I = 1 To 5
        listBoucle(I) = listModalites
    Next
    Dim firstcell As String
    Dim text2analyse As String
    firstcell = Mid(Cla(1, 1), 6)
    sepList = Split(firstcell, "[")
    Dim errorGetExt As Boolean
    errorGetExt = False
    Dim libErrorGetExt As String
    libErrorGetExt = ""
    For I = (LBound(sepList) + 1) To UBound(sepList)
        text2analyse = Split(sepList(I), "]")(0)
        If Left(text2analyse, 2) = "a>" Then
            listBoucle(I) = getExtended("per", True, text2analyse, areaList, 0, "AREA", "", "L")
            If UBound(listBoucle(I)) = 0 Then
                errorGetExt = True
                libErrorGetExt = libErrorGetExt & ";" & text2analyse
            End If
        End If
        If Left(text2analyse, 2) = "s>" Then
            listBoucle(I) = getExtended("per", True, text2analyse, scenarioList, 0, "SCENARIO", "", "L")
            If UBound(listBoucle(I)) = 0 Then
                errorGetExt = True
                libErrorGetExt = libErrorGetExt & ";" & text2analyse
            End If
        End If
        If Left(text2analyse, 2) <> "s>" And Left(text2analyse, 2) <> "a>" Then
            listBoucle(I) = getExtended("per", True, text2analyse, NOMENCLA, 0, "NOMENCLATURE", "", "L")
            If UBound(listBoucle(I)) = 0 Then
                errorGetExt = True
                libErrorGetExt = libErrorGetExt & ";" & text2analyse
            End If
        End If
    Next
    Dim textRec() As String
    ReDim textRec(0)
    If errorGetExt Then
newLogLine = Array("", "ERREUR", nameSheetEnCours, "", time(), Mid(libErrorGetExt, 2), "", "", "Itération incomplète")
logNom = alimLog(logNom, newLogLine)
    nbErrors = 1
GoTo suite
    End If
    
    Dim dimTextRec As Integer
    If UBound(listBoucle(1)) > 0 Then dimTextRec = UBound(listBoucle(1))
    If UBound(listBoucle(2)) > 0 Then dimTextRec = dimTextRec * UBound(listBoucle(2))
    If UBound(listBoucle(3)) > 0 Then dimTextRec = dimTextRec * UBound(listBoucle(3))
    If UBound(listBoucle(4)) > 0 Then dimTextRec = dimTextRec * UBound(listBoucle(4))
    If UBound(listBoucle(5)) > 0 Then dimTextRec = dimTextRec * UBound(listBoucle(5))
    If dimTextRec > 0 Then
        ReDim textRec(1 To dimTextRec)
    End If
    Dim compteur As Integer
    compteur = 0
    If UBound(listBoucle(1)) > 0 Then
        For b1 = LBound(listBoucle(1)) To UBound(listBoucle(1))
            If UBound(listBoucle(2)) > 0 Then
                For b2 = LBound(listBoucle(2)) To UBound(listBoucle(2))
                    If UBound(listBoucle(3)) > 0 Then
                        For b3 = LBound(listBoucle(3)) To UBound(listBoucle(3))
                            If UBound(listBoucle(4)) > 0 Then
                                For b4 = LBound(listBoucle(4)) To UBound(listBoucle(4))
                                    If UBound(listBoucle(5)) > 0 Then
                                        For b5 = LBound(listBoucle(5)) To UBound(listBoucle(5))
                                            compteur = compteur + 1
    textRec(compteur) = "[" & listBoucle(1)(b1) & "]" & "[" & listBoucle(2)(b2) & "]" & "[" & listBoucle(3)(b3) & "]" & "[" & listBoucle(4)(b4) & "]" & "[" & listBoucle(5)(b5) & "]"
                                        Next
                                    Else
                                        compteur = compteur + 1
    textRec(compteur) = "[" & listBoucle(1)(b1) & "]" & "[" & listBoucle(2)(b2) & "]" & "[" & listBoucle(3)(b3) & "]" & "[" & listBoucle(4)(b4) & "]"
                                    End If
                                Next
                            Else
                                compteur = compteur + 1
    textRec(compteur) = "[" & listBoucle(1)(b1) & "]" & "[" & listBoucle(2)(b2) & "]" & "[" & listBoucle(3)(b3) & "]"
                            End If
                        Next
                    Else
                        compteur = compteur + 1
    textRec(compteur) = "[" & listBoucle(1)(b1) & "]" & "[" & listBoucle(2)(b2) & "]"
                    End If
                Next
            Else
                compteur = compteur + 1
    textRec(compteur) = "[" & listBoucle(1)(b1) & "]"
            End If
        Next

    Else
        ' en erreur?
    
    End If
    sepList = Split(firstcell, "]")
    patron = Trim(sepList(UBound(sepList)))
    a1 = ""
    a2 = ""
    a3 = ""
    a4 = ""
    a5 = ""
    p1 = ""
    p2 = ""
    p3 = ""
    p4 = ""
    p5 = ""
    t1 = ""
    t2 = ""
    t3 = ""
    t4 = ""
    t5 = ""
    meta = sheet2process
    patron = Replace(patron, "@META", meta)
    Dim nomFeuille As String
    For f = LBound(textRec) To UBound(textRec)
        nomFeuille = patron
        sepList = Split(textRec(f), "[")
        For s = (LBound(sepList) + 1) To UBound(sepList)
            Chemin = Split(sepList(s), "]")(0)
            If s = 1 Then p1 = Chemin: t1 = Split(Chemin, ">")(UBound(Split(Chemin, ">"))): a1 = p1
            If s = 2 Then p2 = Chemin: t2 = Split(Chemin, ">")(UBound(Split(Chemin, ">"))): a2 = p2
            If s = 3 Then p3 = Chemin: t3 = Split(Chemin, ">")(UBound(Split(Chemin, ">"))): a3 = p3
            If s = 4 Then p4 = Chemin: t4 = Split(Chemin, ">")(UBound(Split(Chemin, ">"))): a4 = p4
            If s = 5 Then p5 = Chemin: t5 = Split(Chemin, ">")(UBound(Split(Chemin, ">"))): a5 = p5
        Next
        nomFeuille = Replace(nomFeuille, "@1", a1)
        nomFeuille = Replace(nomFeuille, "@2", a2)
        nomFeuille = Replace(nomFeuille, "@3", a3)
        nomFeuille = Replace(nomFeuille, "@4", a4)
        nomFeuille = Replace(nomFeuille, "@5", a5)
        nomFeuille = Replace(nomFeuille, "@THIS1", t1)
        nomFeuille = Replace(nomFeuille, "@THIS2", t2)
        nomFeuille = Replace(nomFeuille, "@THIS3", t3)
        nomFeuille = Replace(nomFeuille, "@THIS4", t4)
        nomFeuille = Replace(nomFeuille, "@THIS5", t5)
        nomFeuille = Replace(nomFeuille, "@PATH1", p1)
        nomFeuille = Replace(nomFeuille, "@PATH2", p2)
        nomFeuille = Replace(nomFeuille, "@PATH3", p3)
        nomFeuille = Replace(nomFeuille, "@PATH4", p4)
        nomFeuille = Replace(nomFeuille, "@PATH5", p5)
Application.StatusBar = "Traitement de " & nameSheetEnCours & " : " & nomFeuille
        ' traiter le cas ou a feuille = la feuille meta !!! ici
        Dim indexfeuille As Integer
        If nomFeuille = sheet2process Then
newLogLine = Array("", "ERREUR", nameSheetEnCours, "", time(), nomFeuille, "", "", "Le nom de la feuille dupliquée ne doit pas être celui de la méta feuille")
logNom = alimLog(logNom, newLogLine)
            nbErrors = nbErrors + 1
        Else
            indexfeuille = g_WB_Modele.Worksheets(sheet2process).Index
            If WsExist(nomFeuille, g_WB_Modele) Then
                With Application
                    .ScreenUpdating = False
                    .DisplayAlerts = False
                End With
                indexfeuille = g_WB_Modele.Worksheets(nomFeuille).Index
                g_WB_Modele.Worksheets(nomFeuille).Delete
                Application.ScreenUpdating = True
            End If
            If indexfeuille > 1 Then g_WB_Modele.Worksheets(sheet2process).Copy After:=g_WB_Modele.Worksheets(indexfeuille - 1)
            If indexfeuille = 1 Then g_WB_Modele.Worksheets(sheet2process).Copy Before:=g_WB_Modele.Worksheets(indexfeuille)
            ActiveSheet.NAME = nomFeuille
            ' constitution des feuilles générées par la meta feuille
            Set FLAEC = g_WB_Modele.Worksheets(nomFeuille)
            With FLAEC.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
                ReDim Cla(1 To derligsheet, 1 To DerColSheet)
                Cla = .VALUE
            End With
            For Nolig = 1 To derligsheet
                For NoCol = 1 To DerColSheet
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@1", a1)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@2", a2)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@3", a3)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@4", a4)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@5", a5)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@THIS1", t1)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@THIS2", t2)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@THIS3", t3)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@THIS4", t4)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@THIS5", t5)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@PATH1", p1)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@PATH2", p2)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@PATH3", p3)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@PATH4", p4)
                    Cla(Nolig, NoCol) = Replace(Cla(Nolig, NoCol), "@PATH5", p5)
                Next
            Next
            Cla(1, 1) = "NOP_Col NOP_Row"
            With FLAEC.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
                .VALUE = Cla
            End With
            Dnamesgen = addToList(Dnamesgen, nomFeuille)
newLogLine = Array("", "INFO", nameSheetEnCours, "", time(), nomFeuille, "", "", "")
logNom = alimLog(logNom, newLogLine)
        End If
    Next
    setMetaSheet = Dnamesgen
suite:
    FLCTL.Cells(g_CONTROL_META_CR_L + posFeuille - 1, g_CONTROL_META_CR_C + 1).VALUE = "" & UBound(textRec)
    FLCTL.Cells(g_CONTROL_META_CR_L + posFeuille - 1, g_CONTROL_META_CR_C + 2).VALUE = "" & nbErrors
    sngChrono = Timer - sngChrono
    FLCTL.Cells(g_CONTROL_META_CR_L + posFeuille - 1, g_CONTROL_META_CR_C + 3).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
Exit Function
errorHandler: Call onErrDo("Il y a des erreurs", "setMetaSheet"): Exit Function
errorHandlerWsNotExists: Call onErrDo("La feuille n'existe pas", "setSheet"): Exit Function
End Function
''' Génération d'une feuille étendue à partir d'une feuille générique
Sub setSheet(sheet2process As String)
    On Error GoTo errorHandler
    Dim targetName As String
    Dim sourceNAme As String
    openFileIfNot ("MODELE")
    openFileIfNot ("TARGET")
    Dim FLICI As Worksheet
    Dim FLCIB As Worksheet
    Dim FLCTL As Worksheet
    Dim FLCDG As Worksheet
    Dim FLQUA As Worksheet
    Set FLCTL = Worksheets(g_CONTROL)
    Dim nameSheetEnCours As String
    Dim nameTime As String
    Dim nameArea As String
    Dim nameScenario As String
    nameTime = ActiveSheet.cbTime.VALUE
    nameArea = ActiveSheet.cbArea.VALUE
    nameScenario = ActiveSheet.cbScenario.VALUE
    nameSheetEnCours = sheet2process

    Set FLCDG = setWS(g_CRDATA, g_WB_Modele)
    Set FLCIB = setOrCreateWsfromModel(sheet2process, "CIBLE")
    If Not WsExist(nameSheetEnCours, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLICI = setWS(nameSheetEnCours, g_WB_Modele)
    If Not WsExist(g_QUANTITY, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLQUA = g_WB_Modele.Worksheets(g_QUANTITY)
    If Not WsExist(g_NOMENCLATURE, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLNOM = g_WB_Modele.Worksheets(g_NOMENCLATURE)
    If Not WsExist(g_AREA, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLAREA = g_WB_Modele.Worksheets(g_AREA)
    If Not WsExist(g_SCENARIO, g_WB_Modele) Then GoTo errorHandlerWsNotExists
    Set FLSCENARIO = g_WB_Modele.Worksheets(g_SCENARIO)
    ' Lecture des quantités

    Dim qua() As Variant
    Dim Quantities() As String
    Dim quantities2() As String
    Dim derlig As Integer
    Dim dercol As Integer
Application.StatusBar = "Traitement de " & nameSheetEnCours
    derlig = getDerLig(FLQUA)
    dercol = getDerCol(FLQUA)
    With FLQUA.Range("a" & 2 & ":" & DecAlph(dercol) & derlig)
        ReDim qua(1 To (derlig - 1), 1 To dercol)
        qua = .VALUE
    End With
    Dim quaPos() As String
    ReDim quaPos(1 To UBound(qua, 1), 1 To UBound(qua, 2))
    For I = LBound(qua, 1) To UBound(qua, 1)
        For j = LBound(qua, 2) To UBound(qua, 2)
            quaPos(I, j) = qua(I, j)
            '''quaPos(i, j + 1) = ""
        Next
    Next
    ReDim Quantities(1 To UBound(qua, 1))
    ReDim quantities2(1 To UBound(qua, 1))
    For I = LBound(qua, 1) To UBound(qua, 1)
        Quantities(I) = "[" & qua(I, 3) & "][" & qua(I, 4) & "][" & qua(I, 5) & "]." & qua(I, 6)
        quantities2(I) = I & "@[" & qua(I, 3) & "][" & qua(I, 4) & "][" & qua(I, 5) & "]." & qua(I, 6) & "@"
    Next
    ' fin lecture
    sngChrono = Timer
    Dim LigDeb0No As Integer
    LigDeb0No = 1
    derlig = getDerLig(FLCDG)
    'derlig = Split(FLCDG.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLCDG.Cells(Nolig, 1) <> "" Then LigDebCNo = Nolig + 1
        'If FLCDG.Cells(Nolig, 1) = "CONTEXTE" Then
            'LigDebCNo = Nolig + 1
            'Exit For
        'End If
    Next
    Dim derligsheet As Integer
    Dim DerColSheet As Integer
    derligsheet = getDerLig(FLCDG)
    DerColSheet = getDerCol(FLCDG)
    'DerLigSheet = FLCDG.Range("A" & FLCDG.Rows.Count).End(xlUp).Row
    'DerColSheet = Cells(LigDebCNo - 1, FLCDG.Columns.Count).End(xlToLeft).Column
    '''FLCDG.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, DerLigSheet)).Clear
    With FLCDG.Range("a1:" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Dim newLogLine() As Variant
newLogLine = Array("DATAGEN", "DEBUT", FLICI.NAME, "GENERATION", time())
logNom = alimLog(logNom, newLogLine)
    ' Lecture des paramètres
    '''With FLICI.Range("a1:h1")
        '''param = .VALUE
    '''End With
    'Dim nameTime As String
    'nameTime = ""
    'Dim nameArea As String
    'nameArea = ""
    'Dim nameScenario As String
    'nameScenario = ""
    '''If ActiveSheet.NAME <> "CONTROL" Then
        '''For i = LBound(param, 2) To UBound(param, 2)
            '''If param(1, i) Like "TIME=*" Then nameTime = Split(param(1, i), "=")(1)
            '''If param(1, i) Like "AREA=*" Then nameArea = Split(param(1, i), "=")(1)
            '''If param(1, i) Like "SCENARIO=*" Then nameScenario = Split(param(1, i), "=")(1)
        '''Next
    '''End If
    '''If nameTime = "" Then nameTime = "TIME"
    '''If nameArea = "" Then nameArea = "AREA"
    '''If nameScenario = "" Then nameScenario = "SCENARIO"
    '''Set FLAREA = g_WB_Modele.Worksheets(nameArea)
    '''Set FLSCENARIO = g_WB_Modele.Worksheets(nameScenario)
    ' lecture de TIME
    Dim times() As String
    times = lectureTime(nameTime)
    ' lecture des areas
    Dim AREAS() As Variant
    Dim areaList() As String
    Dim areaName() As String
    Dim ARE() As String
    derlignom = getDerLig(FLAREA)
    'derlignom = FLAREA.Range("A" & FLAREA.Rows.Count).End(xlUp).Row
    AREAS = FLAREA.Range("A2:B" & derlignom).VALUE
    ReDim areaList(1 To UBound(AREAS, 1))
    ReDim areaName(1 To UBound(AREAS, 1))
    ReDim ARE(1 To UBound(AREAS, 1))
    For a = LBound(AREAS, 1) To UBound(AREAS, 1)
        ARE(a) = AREAS(a, 1)
        areaList(a) = AREAS(a, 1) & "@" & AREAS(a, 2)
        areaName(a) = AREAS(a, 2)
    Next
    racineAREA = AREAS(1, 1)
    ' lecture des scénarios
    Dim SCENARIOS() As Variant
    Dim scenarioList() As String
    Dim SCENAR() As String
    Dim scenarioName() As String
    derlignom = getDerLig(FLSCENARIO)
    'derlignom = FLSCENARIO.Range("A" & FLSCENARIO.Rows.Count).End(xlUp).Row
    SCENARIOS = FLSCENARIO.Range("A2:B" & derlignom).VALUE
    ReDim scenarioList(1 To UBound(SCENARIOS, 1))
    ReDim scenarioName(1 To UBound(SCENARIOS, 1))
    ReDim SCENAR(1 To UBound(SCENARIOS, 1))
    For a = LBound(SCENARIOS, 1) To UBound(SCENARIOS, 1)
        SCENAR(a) = SCENARIOS(a, 1)
        scenarioList(a) = SCENARIOS(a, 1) & "@" & SCENARIOS(a, 2)
        scenarioName(a) = SCENARIOS(a, 2)
        'MsgBox scenarioName(a)
    Next
    racineSCENARIO = SCENARIOS(1, 1)
    ' lecture de la nomenclature
    Dim NOMENC() As Variant
    Dim NOMENCLA() As String
    Dim NOMENCLANAME() As String
    Dim IDNOMENC() As Variant
    Dim IDNOMENCLA() As String
    Dim eten() As String
    ReDim eten(0)
    derlignom = FLNOM.Range("C" & FLNOM.Rows.Count).End(xlUp).Row
    NOMENC = FLNOM.Range("c2:e" & derlignom).VALUE
    IDNOMENC = FLNOM.Range("f2:f" & derlignom).VALUE
    ' alimentation des attributs
    na = setATTRIBUTS("NOMENCLATURE")
    '''
    ReDim NOMENCLA(1 To UBound(NOMENC, 1))
    ReDim NOMENCLANAME(1 To UBound(NOMENC, 1))
    ReDim IDNOMENCLA(1 To UBound(IDNOMENC, 1))
    For nol = LBound(NOMENC, 1) To UBound(NOMENC, 1)
        NOMENCLA(nol) = NOMENC(nol, LBound(NOMENC, 2))
        NOMENCLANAME(nol) = NOMENCLA(nol) & "@" & NOMENC(nol, UBound(NOMENC, 2))
        IDNOMENCLA(nol) = IDNOMENC(nol, LBound(IDNOMENC, 2))
    Next
    'Alimentation de Cla tableau de la feuille en cours
    Dim Cla() As Variant
    Dim ClaLig() As Variant
    Dim ClaCol() As Variant
    derligsheet = FLICI.Cells.SpecialCells(xlCellTypeLastCell).Row
    DerColSheet = FLICI.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLICI.Range("a1:a" & derligsheet)
        ReDim ClaLig(1 To derligsheet, 1)
        ClaLig = .VALUE
    End With
    Dim FirLigSheet As Integer
    FirLigSheet = 1
    For Nolig = LBound(ClaLig, 1) To UBound(ClaLig, 1)
        If ClaLig(Nolig, 1) = "BEGIN" Then FirLigSheet = Nolig
        If ClaLig(Nolig, 1) = "END" Then
            derligsheet = Nolig - 1
            Exit For
        End If
    Next
    With FLICI.Range("a1:" & DecAlph(DerColSheet) & "1")
        ReDim ClaCol(1, 1 To DerColSheet)
        ClaCol = .VALUE
    End With
    For NoCol = LBound(ClaCol, 2) To UBound(ClaCol, 2)
        If Trim(ClaCol(1, NoCol)) = "END" Then
            DerColSheet = NoCol - 1
            Exit For
        End If
    Next
    With FLICI.Range("a1:" & DecAlph(DerColSheet) & derligsheet)
        ReDim Cla(1 To derligsheet, 1 To DerColSheet)
        Cla = .VALUE
    End With
    ' tableau résultant
    Dim res() As Variant
    ReDim res(1 To UBound(Cla, 2), 0)
    ' boucle pour étendre en lignes
    Dim debIterLig As Integer
    Dim finIterLig As Integer
    debIterLig = 0
    finIterLig = 0
    Dim numlig As Integer
    Dim listLines() As String
    ReDim listLines(0)
    Dim ent As String
    Dim clalu As String
    Dim norowcollu As String
    Dim numligorcol As Integer
    Dim nbliggener As Integer
    nbliggener = 0
    Dim nbliginit As Integer
    nbliginit = 0
    Dim enti As String
    Dim listD() As String
    Dim nenti As String

    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        If Left(Cla(Nolig, 1), 1) <> "#" Then
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : EXTENSION EN LIGNE : " & Nolig & " / " & UBound(Cla, 1)
            For NoCol = LBound(Cla, 2) To LBound(Cla, 2) ' uniquement 1ere colonne pour l'instant
                If Cla(1, NoCol) Like "*NOP_Col*" Then
                    nbliginit = nbliginit + 1
                    ' Début d'un itération
                    If Left(Cla(Nolig, NoCol), 1) = "(" And Nolig >= FirLigSheet Then
                        debIterLig = Nolig
                        numlig = Nolig
                        Dim iterLignes As Variant
                        iterLignes = iterColOrLig(Trim(Cla(Nolig, NoCol)), numlig, times, NOMENCLA, ARE, SCENAR, FLICI, "L")
'MsgBox Nolig & ":" & LBound(iterLignes) & ":" & UBound(iterLignes) & "=" & Trim(Cla(Nolig, NoCol))
                        If Mid(iterLignes(0)(1), 1, 5) = "ERROR" Then
                            finIterLig = UBound(Cla, 1)
                            GoTo continueLoop
                        End If
                        finIterLig = UBound(Cla, 1)
                    'End If
                    End If
                    ' Traitement de l'itération
'MsgBox Trim(Cla(Nolig, NoCol))
                    If Nolig >= debIterLig And Nolig <= finIterLig Then
                        ' construction des lignes itérées
                        '''lig2process = Split(Split("=" & Trim(Cla(Nolig, NoCol)), "=")(UBound(Split("=" & Trim(Cla(Nolig, NoCol)), "="))), ")")(0)
                        lig2process = Split("=" & Trim(Cla(Nolig, NoCol)), "=")(UBound(Split("=" & Trim(Cla(Nolig, NoCol)), "=")))
                        ''''' pourquoi enlver le dernier caractère ?
                        If Nolig = finIterLig Then lig2process = Mid(lig2process, 1, Len(lig2process) - 1)
                        If UBound(listLines) = 0 Then
                            ReDim listLines(1 To 1)
                        Else
                            ReDim Preserve listLines(1 To UBound(listLines) + 1)
                        End If
                        listLines(UBound(listLines)) = lig2process
'MsgBox Join(listLines, Chr(10)) & Chr(10) & ">>>" & lig2process & "<<<" & Trim(Cla(Nolig, NoCol)) & Chr(10) & Nolig & ":" & finIterLig
                    End If

                    ' Fin d'un itération
                    If Right(Cla(Nolig, NoCol), 1) = ")" And Nolig >= FirLigSheet Then
'MsgBox Nolig & Chr(10) & Trim(Cla(Nolig, NoCol)) & Chr(10) & Join(listLines, Chr(10))
                        finIterLig = Nolig
                        Dim a0() As String
                        Dim a1() As Variant
                        a0 = iterLignes(0)
                        If Mid(iterLignes(0)(1), 1, 5) = "ERROR" Then GoTo continueLoop
                        a1 = iterLignes(1)
'MsgBox numlig & Chr(10) & Cla(Nolig, NoCol)
'MsgBox Join(listLines, Chr(10))
                        listExtended = getMatch(a0, a1, listLines, numlig, NOMENCLA, areaList, scenarioList)
'MsgBox Trim(Cla(Nolig, NoCol)) & Chr(10) & Join(listLines, Chr(10)) & Chr(10) & Chr(10) & Join(listExtended, Chr(10))
                        lastResUbound = UBound(res, 2)
                        ReDim Preserve res(1 To UBound(res, 1), 1 To UBound(res, 2) + UBound(listExtended))
'MsgBox LBound(listExtended) & ":" & UBound(listExtended) & Chr(10) & Join(listExtended, Chr(10))
                        For R = LBound(listExtended) To UBound(listExtended)
                            If R > 0 Then
                                lig = debIterLig + CInt(Split(listExtended(R) & ".", ".")(0)) - 1
    res(1, lastResUbound + R) = lig & "+[" & Split(listExtended(R), "+")(1) & "]." & Split(Split(Split(listExtended(R), "+")(0), "-")(0), ".")(1)
    'Res(1, lastResUbound + r) = lig & "+[" & Split(listExtended(r), "+")(1) & "]." & Split(Split(listExtended(r), "+")(0), ".")(0)
                                If Right(res(1, lastResUbound + R), 1) = "2" Then
                                    If Split(Split(listExtended(R), "+")(0), "_")(1) = "1" Then
                                        res(1, lastResUbound + R) = res(1, lastResUbound + R) & ".1"
                                    End If
                                End If
                                ' ajouté !!!
                                If Right(res(1, lastResUbound + R), 1) = "3" Then
                                    If Split(Split(listExtended(R), "+")(0), "_")(1) = "1" Then
                                        res(1, lastResUbound + R) = res(1, lastResUbound + R) & ".1"
                                    End If
                                End If
'MsgBox listExtended(R) & Chr(10) & res(1, lastResUbound + R)
                                For cc = 2 To UBound(Cla, 2)
                                    norowcollu = res(1, lastResUbound + R)
                                    clalu = Cla(lig, cc)
                                    clalu = epurer(norowcollu, clalu)
                                    res(cc, lastResUbound + R) = clalu
                                Next
                            End If
                        Next
                        ReDim listLines(0)
'MsgBox "fin"
                    End If
                    If (Nolig < debIterLig Or Nolig > finIterLig) And (Nolig >= FirLigSheet Or Nolig = 1) Then
                        ' Résolution périmètre sans itération
                        nenti = Cla(Nolig, 1)
                        If Left(Cla(Nolig, 1), 1) = "[" Then
                            spll = Split(Trim(Cla(Nolig, 1)), "][")
                            numlig = Nolig
                            nenti = ""
                            For s = LBound(spll) To UBound(spll)
                                enti = spll(s)
                                If Left(enti, 1) = "[" Then enti = Mid(enti, 2)
                                If Right(enti, 1) = "]" Then enti = Mid(enti, 1, Len(enti) - 1)
                                If Not enti Like "*$*" Then
                                    listD = NOMENCLA
                                    If Left(enti, 1) = "a" Then listD = ARE
                                    If Left(enti, 1) = "s" Then listD = SCENAR
                                    perimeters = getExtended("feu", True, enti, listD, numlig, "QUANTITE", FLICI.NAME, "L")
                                    If UBound(perimeters) <> 1 Then
'MsgBox nenti & "<<<ERROR>>>" & enti
                                        nenti = nenti & "[" & enti & "]"
                                    Else
                                        nenti = nenti & "[" & perimeters(UBound(perimeters)) & "]"
'MsgBox nenti & "<<<ligne ok>>>" & enti
                                    End If
                                Else
                                    nenti = nenti & "[" & enti & "]"
                                End If
                            Next
'MsgBox ">" & nenti & "<" & Chr(10) & ">" & Cla(Nolig, 1) & "<"
                        End If
                        lastResUbound = UBound(res, 2)
                        If UBound(res, 2) = 0 Then
                            ReDim res(1 To UBound(res, 1), 1 To 1)
                        Else
                            ReDim Preserve res(1 To UBound(res, 1), 1 To UBound(res, 2) + 1)
                        End If
                        For c = LBound(Cla, 2) To UBound(Cla, 2)
                            If c = LBound(Cla, 2) Then
                                'res(c, lastResUbound + 1) = Nolig & "+" & Cla(Nolig, c)
                                res(c, lastResUbound + 1) = Nolig & "+" & nenti
                            Else
                                ent = Cla(Nolig, c)
                                res(c, lastResUbound + 1) = ent
                            End If
                        Next
                    End If
                End If
            Next
        End If
continueLoop:
    Next

    Cla1 = Application.Transpose(res)
    nbliggener = UBound(Cla1, 1)
    'Dim FLCIBTL As Worksheet
    'Set FLCIBTL = Worksheets("TEMPL")
    'FLCIBTL.Cells.Clear
'MsgBox sheet2process & Chr(10) & LBound(Cla, 1) & ":" & UBound(Cla, 1)
'FLCIBTL.Range("a1:" & DecAlph(UBound(Cla1, 2)) & UBound(Cla1, 1)).VALUE = Cla1
    ' boucle pour étendre en colonnes
    ReDim res(1 To UBound(Cla1, 1), 0)
    Dim debIterCol As Integer
    Dim finIterCol As Integer
    Dim numCol As Integer
    Dim listCols() As String
    ReDim listCols(0)
    Dim firstRow As Integer
    firstRow = 0
    Dim nbcolgener As Integer
    nbcolgener = 0
    Dim nbcolinit As Integer
    nbcolinit = 0
    For NoCol = LBound(Cla1, 2) To UBound(Cla1, 2)
        If Left(Cla1(1, NoCol), 1) <> "#" Then
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : EXTENSION EN COLONNE : " & NoCol & " / " & UBound(Cla1, 2)
            For Nolig = LBound(Cla1, 1) To LBound(Cla1, 1) ' uniquement 1 ere ligne pour l'instant
                nbcolinit = nbcolinit + 1
                If Trim(Cla1(Nolig, NoCol)) Like "*NOP_Row*" Then
                    firstRow = firstRow + 1
                End If
                If Trim(Cla1(Nolig, 1)) Like "*NOP_Col*" Then
                    ' Début d'un itération
                    If Left(Cla1(1, NoCol), 1) = "(" Then
                        debIterCol = NoCol
                        numCol = NoCol
                        Dim iterCols As Variant
'MsgBox Nolig & ":ici:" & NoCol & Chr(10) & Trim(Cla1(Nolig, NoCol))
                        iterCols = iterColOrLig(Trim(Cla1(Nolig, NoCol)), numCol, times, NOMENCLA, ARE, SCENAR, FLICI, "C")
'MsgBox Trim(Cla1(Nolig, NoCol)) & Chr(10) & Mid(iterCols(0)(1), 1, 5)
                        If Mid(iterCols(0)(1), 1, 5) = "ERROR" Then GoTo continueLoopCol
                        finIterCol = UBound(Cla1, 2)
                    End If
                    ' Traitement de l'itération
'MsgBox Nolig & ":" & NoCol
                    If NoCol >= debIterCol And NoCol <= finIterCol Then
                        ' construction des colonnes itérées
                        '''Col2process = Split(Split("=" & Trim(Cla1(Nolig, NoCol)), "=")(UBound(Split("=" & Trim(Cla1(Nolig, NoCol)), "="))), ")")(0)
                        Col2process = Split("=" & Trim(Cla1(Nolig, NoCol)), "=")(UBound(Split("=" & Trim(Cla1(Nolig, NoCol)), "=")))
                        If NoCol = finIterCol Then Col2process = Mid(Col2process, 1, Len(Col2process) - 1)
                        If UBound(listCols) = 0 Then
                            ReDim listCols(1 To 1)
                        Else
                            ReDim Preserve listCols(1 To UBound(listCols) + 1)
                        End If
                        listCols(UBound(listCols)) = Col2process
                    End If
                    ' Fin d'un itération
'MsgBox Right(Cla1(Nolig, NoCol), 1)
                    If Right(Cla1(Nolig, NoCol), 1) = ")" Then
                        finIterCol = NoCol
                        Dim c0() As String
                        Dim c1() As Variant
                        c0 = iterCols(0)
                        c1 = iterCols(1)
                        listExtended = getMatch(c0, c1, listCols, numCol, NOMENCLA, areaList, scenarioList)
                        lastResUbound = UBound(res, 2)
'MsgBox "setSheet=" & numCol & Chr(10) & Cla1(Nolig, NoCol) & Chr(10) & UBound(listExtended)
                        ReDim Preserve res(1 To UBound(res, 1), 1 To UBound(res, 2) + UBound(listExtended))
                        For R = LBound(listExtended) To UBound(listExtended)
                            If R > 0 Then
                                lig = debIterCol + CInt(Split(listExtended(R) & ".", ".")(0)) - 1
'MsgBox lig & "<>" & listExtended(R)
    res(1, lastResUbound + R) = lig & "+[" & Split(listExtended(R), "+")(1) & "]." & Split(Split(Split(listExtended(R), "+")(0), "-")(0), ".")(1)
'MsgBox res(1, lastResUbound + R) & Chr(10) & listExtended(R)
                                If Right(res(1, lastResUbound + R), 1) = "2" Then
                                    If Split(Split(listExtended(R), "+")(0), "_")(1) = "1" Then
                                        res(1, lastResUbound + R) = res(1, lastResUbound + R) & ".1"
                                    End If
                                End If
                                For cc = 2 To UBound(Cla1, 1)
                                    norowcollu = res(1, lastResUbound + R)
                                    clalu = Cla1(cc, lig)
                                    clalu = epurer(norowcollu, clalu)
                                    res(cc, lastResUbound + R) = clalu
                                Next
                            End If
                        Next
                        ReDim listCols(0)
                    End If
                    If NoCol < debIterCol Or NoCol > finIterCol Then
                        ' Résolution périmètre sans itération
                        nenti = Cla1(1, NoCol)
                        If Left(Cla1(1, NoCol), 1) = "[" Then
                            spll = Split(Trim(Cla1(1, NoCol)), "][")
                            numCol = NoCol
                            nenti = ""
                            For s = LBound(spll) To UBound(spll)
                                enti = spll(s)
                                If Left(enti, 1) = "[" Then enti = Mid(enti, 2)
                                If Right(enti, 1) = "]" Then enti = Mid(enti, 1, Len(enti) - 1)
                                If Not enti Like "*$*" Then
                                    listD = NOMENCLA
                                    If Left(enti, 1) = "a" Then listD = ARE
                                    If Left(enti, 1) = "s" Then listD = SCENAR
                                    perimeters = getExtended("feu", True, enti, listD, numCol, "QUANTITE", FLICI.NAME, "C")
                                    If UBound(perimeters) <> 1 Then
                                        nenti = nenti & "[" & enti & "]"
                                    Else
                                        nenti = nenti & "[" & perimeters(UBound(perimeters)) & "]"
                                    End If
                                Else
                                    nenti = nenti & "[" & enti & "]"
                                End If
                            Next
                        End If
                        lastResUbound = UBound(res, 2)
                        If UBound(res, 2) = 0 Then
                            ReDim res(1 To UBound(res, 1), 1 To 1)
                        Else
                            ReDim Preserve res(1 To UBound(res, 1), 1 To UBound(res, 2) + 1)
                        End If
                        For c = LBound(Cla1, 1) To UBound(Cla1, 1)
                            If c = LBound(Cla1, 1) And NoCol <> 1 Then
                                'res(c, lastResUbound + 1) = NoCol & "+" & Cla1(c, NoCol)
                                res(c, lastResUbound + 1) = NoCol & "+" & nenti
                            Else
                                ent = Cla1(c, NoCol)
                                res(c, lastResUbound + 1) = ent
                                ' et afficher dans le bon niveau ???
                            End If
                        Next
                    End If
                End If
                'End If
            Next
        End If
continueLoopCol:
    Next
    '''Set FLCIBT1 = Worksheets("TEMP1")
    '''FLCIBT1.Cells.Clear
    '''FLCIBT1.Range("A1:" & DecAlph(UBound(Res, 2)) & UBound(Res, 1)).Value = Res
    'Dim FLCIBTC As Worksheet
    'Set FLCIBTC = Worksheets("TEMPC")
    'FLCIBTC.Cells.Clear

    'FLCIBTC.Range("A1:" & DecAlph(UBound(res, 2)) & UBound(res, 1)).VALUE = res
    nbcolgener = UBound(res, 2)
    ' Résolution des contenus intérieurs aux norow nocol mais que
    Dim cthisn As String
    Dim rthisn As String
    Dim rthis1 As String
    Dim cthis1 As String
    Dim numerolig As Integer
    Dim resnolignocol As String
    Dim perlig As String
    Dim percol As String
    Dim perime As String
    Dim cdn As Integer
    Dim cda As Integer
    Dim cds As Integer
    Dim recent As String
    Dim lastRecent As String
    Dim ligAlerte As Integer
    ligAlerte = 0
    Dim lastLigAlerte As Integer
    lastLigAlerte = 0
    Dim newLigAlerte As Integer
    newLigAlerte = 0
    Dim listOccAlerte As String
    listOccAlerte = ""
    Dim listQuantity() As String
    Dim indexQuantity As Integer
    Dim ctxTime As String
    Dim tim As String
    Dim iqua As Integer
'MsgBox UBound(res, 1) & ":" & UBound(res, 2)
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = True
    reg.Pattern = "\.[^ ]+\([^ ]*\)"
    Dim timeMatch As String
    Dim timqua As String
    compteur = 0
    lastCompteur = 0
    Dim recentctx As String
    Dim ainjecter As String
    Dim sheet2injecter As Integer
    'Dim chaine As String
    'Dim resu() As String
    'Dim resuget() As String
    'Dim resugetboucle As String
    'Dim nu As Integer
    'Dim valeurboucle As String
    'Dim aj As String
    Dim prems As String
    Dim equainit As String
    Dim dateToProcess As String
    Dim debut As Integer
    ' remise à vide des positions de la feuille traitée
    Dim regp As VBScript_RegExp_55.RegExp
    Set regp = New VBScript_RegExp_55.RegExp
    regp.Global = True
    regp.Pattern = "@" & FLCIB.NAME & ";[^@]*@"
    For Nolig = LBound(quaPos, 1) To UBound(quaPos, 1)
        If regp.test(quaPos(Nolig, UBound(qua, 2)) & "@") Then
            amodifier = quaPos(Nolig, UBound(qua, 2)) & "@"
            amodifier = regp.Replace(amodifier, "@")
            quaPos(Nolig, UBound(qua, 2)) = Mid(amodifier, 1, Len(amodifier) - 1)
        End If
    Next
    Dim argstr As String

    For Nolig = LBound(res, 1) To UBound(res, 1)
        numlig = Nolig
        lig1 = res(Nolig, 1)
        For NoCol = LBound(res, 2) To UBound(res, 2)
compteur = Int(100 * ((Nolig - 1) * UBound(res, 2) + NoCol) / (UBound(res, 1) * UBound(res, 2)))
If compteur <> lastCompteur Then
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : RESOLUTION DES CONTENUS : " & ((Nolig - 1) * UBound(res, 2) + NoCol) & " / " & (UBound(res, 1) * UBound(res, 2)) & " " & compteur & " %"
lastCompteur = compteur
DoEvents
End If
'Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : RESOLUTION DES CONTENUS : " & ((Nolig - 1) * UBound(res, 2) + NoCol) & " / " & (UBound(res, 1) * UBound(res, 2))
            col1 = res(1, NoCol)
            lc = Trim(res(Nolig, NoCol))
            If lc Like "*]*" Then
                resnolignocol = res(Nolig, NoCol)
'If NoCol = 1 Then MsgBox lc
                If Nolig > 1 And NoCol > 1 Then
                    percol = Split(res(1, NoCol), "+")(1)
                    perlig = Split(res(Nolig, 1), "+")(1)
'MsgBox ">>>" & res(Nolig, 1) & Chr(10) & res(1, NoCol)
                    If perlig <> "" Then
                        perime = perlig
                    Else
                        perime = percol
                    End If
                    perime = Split(perime, ".1")(0)
                    perime = Split(perime, ".2")(0)
                    perime = Split(perime, ".3")(0)
                    For I = 1 To 4
                        perime = Replace(perime, "THIS" & I & "=", "")
                    Next
                    perime = Replace(perime, "THIS=", "")
                    resnolignocol = Replace(resnolignocol, "[]", perime)
'old = res(Nolig, NoCol)

                    res(Nolig, NoCol) = resnolignocol
'If NoCol = 2 Then MsgBox Nolig & ":" & NoCol & Chr(10) & Trim(res(Nolig, 1)) & Chr(10) & old & Chr(10) & res(Nolig, NoCol)
                    For I = 1 To 4
                        If I = 4 Then
                            cthisn = "CTHIS"
                            cthis1 = "CTHIS1"
                        Else
                            cthisn = "CTHIS" & I
                            cthis1 = "CTHIS" & I
                        End If
                        If res(Nolig, NoCol) Like "*" & cthisn & "*" Then
                            newcthisS = Split(col1, Mid(cthis1, 2) & "=")
                            If UBound(newcthisS) > 0 Then
                                '''newcthis = Split(Split(newcthisS(1), ":")(0), ".")(0)
                                newcthis = Split(Split(newcthisS(1), ".")(0), "]")(0)
                                res(Nolig, NoCol) = Replace(res(Nolig, NoCol), cthisn, newcthis)
                            'ent = Res(Nolig, NoCol)
                            'Res(Nolig, NoCol) = getName(ent, NOMENCLANAME, areaName, scenarioName)
                            End If
                        End If
                        If I = 4 Then
                            rthisn = "RTHIS"
                            rthis1 = "RTHIS1"
                        Else
                            rthisn = "RTHIS" & I
                            rthis1 = "RTHIS" & I
                        End If
                        If res(Nolig, NoCol) Like "*" & rthisn & "*" Then
                            newrthisS = Split(lig1, Mid(rthis1, 2) & "=")
                            If UBound(newrthisS) > 0 Then
                                '''newrthis = Split(Split(Split(newrthisS(1), ":")(0), ".")(0), "]")(0)
                                newrthis = Split(Split(newrthisS(1), ".")(0), "]")(0)
                                res(Nolig, NoCol) = Replace(res(Nolig, NoCol), rthisn, newrthis)
                            'ent = Res(Nolig, NoCol)
                            'Res(Nolig, NoCol) = getName(ent, NOMENCLANAME, areaList, scenarioList)
                            End If
                        End If
                    Next
                    ent = res(Nolig, NoCol)
'MsgBox res(Nolig, 1)
'MsgBox ent & ":" & reg.test(ent) & ":" & reg.Pattern
                    If ent <> "" And InStr(ent, ".") > 0 And InStr(ent, "(") > 0 And InStr(ent, ")") > 0 And reg.test(ent) Then
                        ' résolution des >>
                        ent = remplaceDoubleSup(ent, NOMENCLA, "", 0)
                        res(Nolig, NoCol) = ent
                        splent = Split(ent, "[")
                        cdn = 0
                        cda = 0
                        cds = 0
'MsgBox res(1, NoCol)
                        For I = LBound(splent) + 1 To UBound(splent)
                            aTraiter = splent(I)
                            If Left(aTraiter & ">", 2) = "a>" Then cda = 1
                            If Left(aTraiter & ">", 2) = "s>" Then cds = 1
                            If Left(aTraiter & ">", 2) <> "a>" And Left(aTraiter & ">", 2) <> "s>" Then cdn = 1
                        Next
                        If cda = 0 Then
                            pl = res(Nolig, 1)
                            da = ""
                            If InStr(pl, "[a>") > 0 Then
                                da = "[a>" & Split(Split(Split(pl, "[a>")(1), ".")(0), "]")(0) & "]"
                            End If
                            If da = "" Then
                                pl = Replace(res(1, NoCol), "][", ".")
                                da = ""
                                If InStr(pl, "[a>") > 0 Or InStr(pl, ".a>") > 0 Then
                                    da = "[a>" & Split(Split(Split(pl, "[a>")(1), ".")(0), "]")(0) & "]"
                                End If
                                If InStr(pl, ".a>") > 0 Then
                                    da = "[a>" & Split(Split(Split(pl, "[a>")(1), ".")(0), "]")(0) & "]"
                                End If
                            End If
                            If da = "" Then da = "[a]"
                            da = Replace(da, "THIS1=", "")
                            da = Replace(da, "THIS2=", "")
                            da = Replace(da, "THIS3=", "")
                        End If
'MsgBox da
                        If cds = 0 Then
                            pl = res(Nolig, 1)
                            ds = ""
                            If InStr(pl, "[s>") > 0 Or InStr(pl, ".s>") > 0 Then
                                ds = "[s>" & Split(Split(Split(pl, "[s>")(1), ".")(0), "]")(0) & "]"
                            End If
                            If ds = "" Then
                                pl = Replace(res(1, NoCol), "][", ".")
'MsgBox pl
                                ds = ""
                                If InStr(pl, "[s>") > 0 Then
                                    ds = "[s>" & Split(Split(Split(pl, "[s>")(1), ".")(0), "]")(0) & "]"
                                End If
                                If InStr(pl, ".s>") > 0 Then
'MsgBox Split(pl, ".s>")(1)
                                    ds = "[s>" & Split(Split(Split(pl, ".s>")(1), ".")(0), "]")(0) & "]"
                                End If
                            End If
'MsgBox ds
                            If ds = "" Then ds = "[s]"
                            ds = Replace(ds, "THIS1=", "")
                            ds = Replace(ds, "THIS2=", "")
                            ds = Replace(ds, "THIS3=", "")
                        End If
'MsgBox ds & Chr(10) & recent & Chr(10) & ent
                        If da & ds = "" Then
                            recent = ent
                        Else
                            If UBound(Split(ent, "].")) < 1 Then
                                'ERROR
                                recent = Split(ent, "].")(0) & "]" & da & ds & "." & ""
newLogLine = Array("", "ERREUR", FLICI.NAME, "Quantité", time(), "", recent, "" & Nolig, "Périmètre inconnu")
logNom = alimLog(logNom, newLogLine)
                            Else
                                ' ajout des contextes Area et Scénario ssi ils n'y sont pas
                                entToTransf = ent
                                If InStr(entToTransf, "[a]") = 0 And InStr(entToTransf, "[a>") = 0 Then
                                    ' On ajoute le contexte de la dimension Area
                                    If InStr(entToTransf, "][") > 0 Then
                                        ' cas où il y a un contexte Scénario
                                        entToTransf = Replace(entToTransf, "][", "]" & da & "[")
                                    Else
                                        entToTransf = Replace(entToTransf, "].", "]" & da & ".")
                                    End If
                                End If
                                If InStr(entToTransf, "[s]") = 0 And InStr(entToTransf, "[s>") = 0 Then
                                    ' On ajoute le contexte de la dimension Scenario
                                    entToTransf = Replace(entToTransf, "].", "]" & ds & ".")
                                End If
                                recent = entToTransf
                                'recent = Split(ent, "].")(0) & "]" & da & ds & "." & Split(ent, "].")(1)
                            End If
                        End If
'MsgBox recent
                        '-- Traitement du temps
                            ' détermination du contexte temps ATTENTION pas de symétrie ligne colonne
                        percol = Replace(percol, "][", ".")
                        splPerCol = Split(percol, ".")
                        For I = LBound(splPerCol) To UBound(splPerCol)
                            If splPerCol(I) Like "*$*" Then
                                ctxTime = Split(splPerCol(I), "]")(0)
                                If Left(ctxTime, 1) = "[" Then ctxTime = Mid(ctxTime, 2)
                                ctxTime = Replace(ctxTime, "THIS1=", "")
                                ctxTime = Replace(ctxTime, "THIS2=", "")
                                ctxTime = Replace(ctxTime, "THIS3=", "")
                                Exit For
                            End If
                        Next
                            ' match entre temps de l'occurence et temps du contexte
                            ' ne présage pas du domaine de définition du temps de la quantité
                        tim = Split(Split(recent & "(", "(")(1) & ")", ")")(0)
                        timeMatch = getTimeMatch(tim, ctxTime, times)
'If InStr(tim, "first") > 0 Then
'MsgBox recent & Chr(10) & timeMatch & ":" & tim & ":" & ctxTime
                        '--
                        If timeMatch = "" Then
                            res(Nolig, NoCol) = ""
newLogLine = Array("", "ALERTE", FLICI.NAME, "Quantité", time(), "", recent, "" & Nolig, "Contexte TIME inconnu")
logNom = alimLog(logNom, newLogLine)
                        Else
                            '''listQuantity = isAquantity2(recent, Quantities2)
                            spl1 = Split(recent, "[")
                            spl2 = Split(spl1(1), "]")
                            prems = spl2(0)
                            ext = getExtended("feu", True, prems, NOMENCLA, numlig, "QUANTITE", FLICI.NAME, "L")
                            If UBound(ext) > 0 Then
                                recent = Replace(recent, "[" & prems & "]", "[" & ext(1) & "]")
                            End If
                            listQuantity = Filter(quantities2, Split(recent, "(")(0) & "@", True)
'MsgBox recent & Chr(10) & UBound(listQuantity)
                            If UBound(listQuantity) > -1 Then
                                ' on va chercher l'équation et les formules après car appel d'une quantité possiblement positionnée après la ligne
                                iqua = 0
'If InStr(LCase(recent), "sommeproduit") > 0 Or InStr(LCase(recent), "sommeprod") > 0 Then
    'MsgBox recent & Chr(10) & Join(listQuantity, Chr(10))
'End If
                                For l = LBound(listQuantity) To UBound(listQuantity)
                                    ' on prend la 1ere quantité dont le domaine time match avec le contexte
                                    ' il faut prendre le dernier qui match plutôt ???
                                    indexQuantity = CInt(Split(listQuantity(l), "@")(0))
                                    timqua = quaPos(indexQuantity, 7)
                                    timocc = Split(timeMatch, ">")(0)
                                    ' détermination si le time de l'occurence match avec le time du contexte
'MsgBox recent & Chr(10) & timqua & ">" & l & "/" & UBound(listQuantity) & ">>" & timeMatch & "<<>>" & dateInDate(timeMatch, timqua, times) & "<<"
                                    If dateInDate(timeMatch, timqua, times) <> "" Then
'MsgBox dateInDate(timeMatch, timqua, times)
                                        iqua = CInt(Split(listQuantity(l), "@")(0))
                                        Exit For
                                    End If
                                Next
                                If iqua > 0 Then
                                    ' a revoir dans le cas d'ubiquïté

                                    splqp = Split(quaPos(iqua, UBound(qua, 2)), "@")
'MsgBox listQuantity(l)
                                    If debut = 0 Or quaPos(iqua, UBound(qua, 2)) = "" Or Not quaPos(iqua, UBound(qua, 2)) Like "*@" & FLCIB.NAME & ";" & "*" Then
                                        ante = FLCIB.NAME & ";"
                                        debut = 1
'MsgBox Nolig & ":new:" & ante
                                    End If
                                        ' chercher la liste à alimenter
                                    sheet2injecter = -1
                                        For l = LBound(splqp) To UBound(splqp)
                                            If splqp(l) Like FLCIB.NAME & ";" & "*" Then
                                                'ante = quaPos(iqua, UBound(qua, 2)) & ";"
                                                'If debut = 1 Then
                                                ante = splqp(l) & ";"
'MsgBox Nolig & ":old:" & ante
                                                'Else
                                                    'ante = FLCIB.NAME & ";"
                                                    'debut = 1
                                                'End If
                                                sheet2injecter = l
                                                Exit For
                                            End If
                                        Next
                                    'End If
                                    ' on extrait le temps du contexte (ATTENTION PAS DE SYMETRIE LIGNE COLONNE)
                                    pl = res(1, NoCol)
                                    spl = Split(pl, ".")
                                    ctx = ""
                                    For I = LBound(spl) To UBound(spl)
                                        If spl(I) Like "*$*" Then
                                            ctx = Split("[" & Split(spl(I), "]")(0), "[")(UBound(Split("[" & Split(spl(I), "]")(0), "[")))
                                            Exit For
                                        End If
                                    Next
                                    ctx = Replace(ctx, "THIS1=", "")
                                    ctx = Replace(ctx, "THIS2=", "")
                                    ctx = Replace(ctx, "THIS3=", "")
                                    ' on reinjecte les positions au bon endroit
                                    ainjecter = ""
                                    If quaPos(iqua, UBound(qua, 2)) = "" Or Not quaPos(iqua, UBound(qua, 2)) Like "*@" & FLCIB.NAME & ";" & "*" Then
                                        'If quaPos(iqua, UBound(qua, 2)) = "" Then
                                            'jct = ""
                                        'Else
                                            'jct = "@"
                                        'End If
                                        ' ça arrive toujours!!!
                                        jct = "@"
                                        quaPos(iqua, UBound(quaPos, 2)) = quaPos(iqua, UBound(qua, 2)) & jct & ante & ctx & ":" & (Nolig - 1) & "," & (NoCol - 1)
'MsgBox "0>" & quaPos(iqua, UBound(quaPos, 2))
'MsgBox "new=" & Nolig & Chr(10) & quaPos(iqua, UBound(quaPos, 2))
                                    Else
'MsgBox Join(splqp, Chr(10))
                                        ' Ecrasement des anciennes coordonnées ??
                                        'ante = ""
                                        For l = LBound(splqp) To UBound(splqp)
                                            If l = LBound(splqp) Then
                                                jct = ""
                                            Else
                                                jct = "@"
                                            End If
                                            'jct = "@"
                                            plus = splqp(l)
                                            If l = sheet2injecter Then
                                                plus = ante & ctx & ":" & (Nolig - 1) & "," & (NoCol - 1)
                                            End If
                                            ainjecter = ainjecter & jct & plus
                                        Next
'MsgBox Nolig & Chr(10) & sheet2injecter & ":ainjecter=" & ainjecter & Chr(10) & "ante=" & ante
                                        '''quaPos(iqua, UBound(quaPos, 2)) = ante & ctx & ":" & (Nolig - 1) & "," & (NoCol - 1)
    'MsgBox ainjecter
                                        quaPos(iqua, UBound(quaPos, 2)) = ainjecter
'MsgBox "1>" & quaPos(iqua, UBound(quaPos, 2))
                                    End If
                                    ' reste à propager correctement le contexte dans les termes de l'équation
                                    ' ou dans le time de la quantité
                                    recentctx = recent
                                    recentctx = Replace(recent, "(" & tim, "(" & timeMatch)
        
                                    equainit = Replace(Trim(quaPos(iqua, 8)), "(t", "(" & ctxTime)
                                    equainit = Replace(equainit, ":t)", ":" & ctxTime & ")")
                                    equainit = Replace(equainit, "(t)", "(" & ctxTime & ")")
                                    equainit = Replace(equainit, "()", "(" & ctxTime & ")")
'If InStr(equainit, "sommep") > 0 Then MsgBox equainit
                                    listermes = analyseTerme(equainit)
                                    For lt = LBound(listermes) To UBound(listermes)
'If InStr(LCase(equainit), "sommeproduit") > 0 Or InStr(LCase(equainit), "sommeprod") > 0 Then
    'MsgBox listermes(lt)
'End If
                                        spltime = Split(Split(listermes(lt) & "(", "(")(1) & ")", ")")(0)
                                        If InStr(spltime, ":") > 0 Then
                                            spltimeplus = Split(spltime, ":")
                                            dateToProcess = spltimeplus(0)
                                            nd = "$" & Split(dateInDate(dateToProcess, timqua, times), "$")(1)
                                            spltime = Replace(spltime, dateToProcess, nd)
                                            dateToProcess = spltimeplus(1)
                                            nd = "$" & Split(dateInDate(dateToProcess, timqua, times), "$")(1)
                                            spltime = Replace(spltime, dateToProcess, nd)
                                            spltimeplus = Split(spltime, ":")
                                            dateToProcess = spltimeplus(0)
                                            If UBound(Split(dateToProcess, ">")) > 0 Then sepd = ">"
                                            If UBound(Split(dateToProcess, ".")) > 0 Then sepd = "."
                                            numdate0 = CInt(Split(dateToProcess, sepd)(1))
                                            typed0 = Split(dateToProcess, sepd)(0)
                                            dateToProcess = spltimeplus(1)
                                            If UBound(Split(dateToProcess, ">")) > 0 Then sepd = ">"
                                            If UBound(Split(dateToProcess, ".")) > 0 Then sepd = "."
                                            numdate1 = CInt(Split(dateToProcess, sepd)(1))
                                            typed1 = Split(dateToProcess, sepd)(0)
                                            corps = Split(listermes(lt) & "(", "(")(0)
                                            newString = ""
                                            For ci = numdate0 To numdate1
                                                If ci = numdate0 Then
                                                    seppv = ""
                                                Else
                                                    seppv = ";"
                                                End If
                                                newString = newString & seppv & corps & "(" & typed0 & ">" & ci & ")"
                                            Next
                                            equainit = Replace(equainit, listermes(lt), newString)
                                        End If
                                    Next
                                    'MsgBox ">>" & equainit & Chr(10) & Join(ltermes, Chr(10))
                                   'a = dateInDate(timeMatch, timqua, times)
'MsgBox ctxTime & ";" & timqua & ":" & timeMatch & "::" & equainit
'MsgBox Trim(quaPos(iqua, 8)) & Chr(10) & equainit & Chr(10) & recentctx & Chr(10) & recent & Chr(10) & ctxTime
'If Right(recent, 2) = ".I" Then MsgBox Right(recent, 2) & Chr(10) & recent & Chr(10) & Trim(quaPos(iqua, 16)) & Left(Trim(quaPos(iqua, 16)), 1)
                                    If (Right(recent, 2) = ".I" Or Right(recent, 2) = ".S") And Left(Trim(quaPos(iqua, 16)), 1) <> "@" Then
                                        res(Nolig, NoCol) = "I@"
                                    Else
                                        ' traitement des calculs complémentaires
                                        'If InStr(LCase(equainit), "sommeproduit") > 0 Or InStr(LCase(equainit), "sommeprod") > 0 Then
'MsgBox equainit
                                        
                                        'End If
                                        If LCase(equainit) = "interpolation" Then
                                            If Left(Right(recentctx, 2), 1) = "." Then
                                                recentctxnew = Left(recentctx, Len(recentctx) - 2)
                                            Else
                                                recentctxnew = recentctx
                                            End If
                                            spld = Trim(Split(Split(Split(recentctxnew, "(")(1), ")")(0), ">")(1))
                                            spldi = CInt(spld) - 1
                                            recentctxnew = Replace(recentctxnew, spld, "" & spldi)
                                            equainit = recentctxnew + " + interpolation"
                    'msgbox equainit
                                        End If
                                        lastequainit = equainit
                                        If InStr(recentctx, ")") <> Len(recentctx) And equainit <> "" Then
'MsgBox Trim(quaPos(iqua, 8)) & Chr(10) & equainit & Chr(10) & d
                                            adroite = Split(recentctx, ")")(1)
                                            lastequainit = equainit
                                            equainit = "(" & equainit & ")" & adroite
                                        End If
                                        If Right(recent, 2) = ".M" Then
                                            ' cas mixte input et équation
                                            If equainit = "" Then
                                                res(Nolig, NoCol) = "I@"
                                            Else
                                                'res(Nolig, NoCol) = "E@" & equainit
                                                res(Nolig, NoCol) = "E@" & lastequainit
                                            End If
                                        Else
                                            If equainit = "" Then
                                                res(Nolig, NoCol) = "E@" & recentctx
                                            Else
                                                res(Nolig, NoCol) = "E@" & equainit
                                            End If
                                        End If
                                    End If
                                    'If quaPos(iqua, 16) <> "" Then
                                        'calcul de la liste
                                        'splQ = Split(quaPos(iqua, 3), ".")
                                        'lastc = Right(quaPos(iqua, 16), 1)
                                        'If lastc = "T" Then lastc = "1"
                                        'ReDim resu(LBound(splQ) To UBound(splQ))
                                        'For t = LBound(splQ) To UBound(splQ)
                                            'If lastc = "" & (t + 1) Then
                                                'chaine = Mid(splQ(t), 1, Len(splQ(t)) - InStr(StrReverse(splQ(t)), ">")) & ">EACH"
                                            'Else
                                                'chaine = splQ(t)
                                            'End If
                                            'resu(t) = chaine
                                        'Next
                                        'chaine = Join(resu, ".")
                                        'resuget = getExtended("feu", True, chaine, listD, numlig, "QUANTITE", FLICI.NAME, "L")
                                        'If UBound(resuget) = 0 Then
                                            ' ERROR
                                        'Else
                                            'nu = -1
                                            'If InStr(quaPos(iqua, 16), "LAST") > 0 Then nu = UBound(resuget)
                                            'If InStr(quaPos(iqua, 16), "FIRST") > 0 Then nu = LBound(resuget)
                                            'If nu = -1 Then
                                                ' ERROR
                                            'Else
                                                'If quaPos(iqua, 3) = resuget(nu) Then
                                                    ' calcul de la nouvelle équation
                                                    'If Trim(quaPos(iqua, 15)) <> "" Then
                                                        'valeurboucle = quaPos(iqua, 15)
                                                    'Else
                                                        'valeurboucle = "1"
                                                    'End If
                                                    'aj = "]" & da & ds & "." & quaPos(iqua, 6) & "(" & timeMatch & ");"
                                                    'resugetboucle = "[" & Join(resuget, aj & "[") & aj
                                                    'resugetboucle = Replace(resugetboucle, "[" & resuget(nu) & aj, "")
                                                    'resugetboucle = valeurboucle & "-somme(" & Mid(resugetboucle, 1, Len(resugetboucle) - 1) & ")"
                                                    'res(Nolig, NoCol) = "E@" & resugetboucle
   ' MsgBox nu & Chr(10) & quaPos(iqua, 3) & Chr(10) & Chr(10) & resugetboucle
                                                'End If
                                            'End If
                                        'End If
'MsgBox UBound(aaa) & Chr(10) & quaPos(iqua, 16) & ":" & lastc & Chr(10) & quaPos(iqua, 3)
'''If UBound(aaa) > 0 Then
    'MsgBox quaPos(iqua, 16) & Chr(10) & quaPos(iqua, 3) & Join(aaa, Chr(10))
'''Else
    'MsgBox quaPos(iqua, 16) & Chr(10) & quaPos(iqua, 3)
'''End If
                                    'End If
                                End If
                            End If
                            'End If
                        End If
                        If iqua = 0 And timeMatch <> "" And Left(res(Nolig, NoCol), 2) <> "I@" And Left(res(Nolig, NoCol), 2) <> "S@" Then
'If Nolig = 5 And NoCol = 3 Then MsgBox res(Nolig, NoCol)
                            res(Nolig, NoCol) = ""
                            newLigAlerte = CInt(Split(res(Nolig, 1), "+")(0))
                            If newLigAlerte <> lastLigAlerte Or recent <> lastRecent Or (Nolig = UBound(res, 1) And NoCol = UBound(res, 2)) Then
newLogLine = Array("", "ALERTE", FLICI.NAME, "Quantité", time(), "", recent, "" & newLigAlerte, "Quantité inconnue")
logNom = alimLog(logNom, newLogLine)
                                lastRecent = recent
                                lastLigAlerte = newLigAlerte
                            End If
                            ligAlerte = CInt(Split(res(Nolig, 1), "+")(0))
                            listOccAlerte = recent
                        End If
                    Else
                        res(Nolig, NoCol) = getName(ent, NOMENCLANAME, areaList, scenarioList)
                    End If
                End If
            End If
            'If NoCol = UBound(Res, 2) Then
                'If ligAlerte = Nolig Then
'newLogLine = Array("", "ALERTE", FLICI.NAME, "Quantité", Time(), "", listOccAlerte, "" & Nolig, "Quantité inconnue")
'logNom = alimLog(logNom, newLogLine)
                'End If
                'ligAlerte = 0
                'listColAlerte = ""
            'End If
            argstr = "" & res(Nolig, NoCol)
            res(Nolig, NoCol) = resolveFunctions(argstr)
        Next
    Next
    '''Set FLCIBT = Worksheets("TEMP")
    'Set FLCIBT = Worksheets(Mid(nameSheetEnCours, 2))
    FLCIB.Cells.Clear
    FLCIB.Range("A1:" & DecAlph(UBound(res, 2)) & UBound(res, 1)).VALUE = res
    ' Constitution des tableaux de formats
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : CONSTITUTION DES FORMATS"
    Dim colWidth() As Double
    Dim ligHeight() As Double
    Dim InteriorColor() As Variant
    Dim FontBold() As Variant
    Dim BorderXlEdgeLeftWeight() As Variant
    Dim BorderXlEdgeLeftLineStyle() As Variant
    Dim BorderXlEdgeRightWeight() As Variant
    Dim BorderXlEdgeRightLineStyle() As Variant
    Dim BorderXlEdgeTopWeight() As Variant
    Dim BorderXlEdgeTopLineStyle() As Variant
    Dim BorderXlEdgeBottomWeight() As Variant
    Dim BorderXlEdgeBottomLineStyle() As Variant
    Dim FontFontStyle() As Variant
    Dim FontName() As Variant
    Dim FontItalic() As Variant
    Dim FontUnderline() As Variant
    Dim FontColor() As Variant
    Dim FontSize() As Variant
    Dim NumberFormat() As Variant
    ReDim colWidth(LBound(Cla, 2) To UBound(Cla, 2))
    ReDim ligHeight(LBound(Cla, 1) To UBound(Cla, 1))
    ReDim InteriorColor(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontBold(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeLeftWeight(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeLeftLineStyle(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeRightWeight(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeRightLineStyle(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeTopWeight(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeTopLineStyle(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeBottomWeight(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim BorderXlEdgeBottomLineStyle(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontFontStyle(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontName(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontItalic(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontUnderline(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontColor(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim FontSize(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ReDim NumberFormat(LBound(Cla, 1) To UBound(Cla, 1), LBound(Cla, 2) To UBound(Cla, 2))
    ompteur = 0
    lastCompteur = 0
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        ligHeight(Nolig) = FLICI.Rows(Nolig).RowHeight
        For NoCol = LBound(Cla, 2) To UBound(Cla, 2)
compteur = Int(100 * ((Nolig - 1) * UBound(res, 2) + NoCol) / (UBound(res, 1) * UBound(res, 2)))
If compteur <> lastCompteur Then
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : COPIE DES FORMATS : " & ((Nolig - 1) * UBound(res, 2) + NoCol) & " / " & (UBound(res, 1) * UBound(res, 2)) & " " & compteur & " %"
lastCompteur = compteur
DoEvents
End If
'Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : COPIE DES FORMATS : " & ((Nolig - 1) * UBound(Cla, 2) + NoCol) & " / " & (UBound(Cla, 1) * UBound(Cla, 2))
            If Nolig = 1 Then colWidth(NoCol) = FLICI.Columns(NoCol).ColumnWidth
            With FLICI.Cells(Nolig, NoCol)
                InteriorColor(Nolig, NoCol) = .Interior.Color
                FontBold(Nolig, NoCol) = .Font.Bold
                BorderXlEdgeLeftWeight(Nolig, NoCol) = .Borders(xlEdgeLeft).Weight
                BorderXlEdgeLeftLineStyle(Nolig, NoCol) = .Borders(xlEdgeLeft).LineStyle
                BorderXlEdgeRightWeight(Nolig, NoCol) = .Borders(xlEdgeRight).Weight
                BorderXlEdgeRightLineStyle(Nolig, NoCol) = .Borders(xlEdgeRight).LineStyle
                BorderXlEdgeTopWeight(Nolig, NoCol) = .Borders(xlEdgeTop).Weight
                BorderXlEdgeTopLineStyle(Nolig, NoCol) = .Borders(xlEdgeTop).LineStyle
                BorderXlEdgeBottomWeight(Nolig, NoCol) = .Borders(xlEdgeBottom).Weight
                BorderXlEdgeBottomLineStyle(Nolig, NoCol) = .Borders(xlEdgeBottom).LineStyle
                FontFontStyle(Nolig, NoCol) = .Font.FontStyle
                FontName(Nolig, NoCol) = .Font.NAME
                FontItalic(Nolig, NoCol) = .Font.Italic
                FontUnderline(Nolig, NoCol) = .Font.Underline
                FontColor(Nolig, NoCol) = .Font.Color
                FontSize(Nolig, NoCol) = .Font.Size
                NumberFormat(Nolig, NoCol) = .NumberFormat
            End With
            '''Next
        Next
    Next
    compteur = 0
    lastCompteur = 0
    
    For Nolig = LBound(res, 1) To UBound(res, 1)
        lig = CInt(Split(res(Nolig, 1), "+")(0))
        FLCIB.Rows(Nolig).RowHeight = ligHeight(lig)
        For NoCol = LBound(res, 2) To UBound(res, 2)
'Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : COLLAGE DES FORMATS : " & ((Nolig - 1) * UBound(res, 2) + NoCol) & " / " & (UBound(res, 1) * UBound(res, 2))
compteur = Int(100 * ((Nolig - 1) * UBound(res, 2) + NoCol) / (UBound(res, 1) * UBound(res, 2)))
If compteur <> lastCompteur Then
Application.StatusBar = "TRAITEMENT DE " & nameSheetEnCours & " : COLLAGE DES FORMATS : " & ((Nolig - 1) * UBound(res, 2) + NoCol) & " / " & (UBound(res, 1) * UBound(res, 2)) & " " & compteur & " %"
lastCompteur = compteur
End If
            col = CInt(Split(res(1, NoCol), "+")(0))
            If Nolig = 1 Then FLCIB.Columns(NoCol).ColumnWidth = colWidth(col)
            FLCIB.Cells(Nolig, NoCol).Interior.Color = InteriorColor(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.Bold = FontBold(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeLeft).Weight = BorderXlEdgeLeftWeight(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeLeft).LineStyle = BorderXlEdgeLeftLineStyle(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeRight).Weight = BorderXlEdgeRightWeight(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeRight).LineStyle = BorderXlEdgeRightLineStyle(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeTop).Weight = BorderXlEdgeTopWeight(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeTop).LineStyle = BorderXlEdgeTopLineStyle(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeBottom).Weight = BorderXlEdgeBottomWeight(lig, col)
            'FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeBottom).LineStyle = BorderXlEdgeBottomLineStyle(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.FontStyle = FontFontStyle(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.NAME = FontName(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.Italic = FontItalic(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.Underline = FontUnderline(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.Color = FontColor(lig, col)
            FLCIB.Cells(Nolig, NoCol).Font.Size = FontSize(lig, col)
            FLCIB.Cells(Nolig, NoCol).NumberFormat = NumberFormat(lig, col)
            'If Nolig = 1 And NoCol = 1 Then
                'MsgBox lig & ":" & col
            'End If
            
        Next
    Next
    ' Traitement des bordures FAUT BOUCLER EN MEME TEMPS COL ET LIG
    Dim lastLig As Integer
    'Dim lig As Integer
    Dim ligne As String
    lastLig = 0
    firslig = 1
    lal = 1
    fil = 1
    For Nolig = LBound(res, 1) To UBound(res, 1)
        ligne = res(Nolig, 1)
        For NoCol = LBound(res, 2) To UBound(res, 2)
            colonne = res(1, NoCol)
            lig = CInt(Split(ligne, "+")(0))
            col = CInt(Split(colonne, "+")(0))
            'if lig <>
        'If (lastLig <> lig And Nolig > 1) Or Nolig = UBound(res, 1) Then
            'If Nolig = UBound(res, 1) Then
                'lal = Nolig
            'Else
                'lal = Nolig - 1
            'End If
            If lastLig <> lig Then
'MsgBox Nolig & ":" & NoCol & Chr(10) & lig & ":" & col & Chr(10) & BorderXlEdgeTopWeight(lig, col) & Chr(10) & BorderXlEdgeTopLineStyle(lig, col)
FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeTop).Weight = BorderXlEdgeTopWeight(lig, col)
FLCIB.Cells(Nolig, NoCol).Borders(xlEdgeTop).LineStyle = BorderXlEdgeTopLineStyle(lig, col)
            'fil = Nolig
            'firslig = lig
            End If
        'End If
        Next
        lastLig = lig
    Next
    FLCIB.Columns(1).EntireColumn.Delete
    FLCIB.Rows(1).EntireRow.Delete
    sngChrono = Timer - sngChrono
    newLogLine = Array("", "INFO", FLICI.NAME, "Statistiques", time(), "", "", nbliginit & " lignes génériques", nbliggener & " lignes générées")
logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("", "INFO", FLICI.NAME, "Statistiques", time(), "", "", nbcolinit & " colonnes génériques", nbcolgener & " colonnes générées")
logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("DATAGEN", "FIN", FLICI.NAME, "GENERATION", time(), "", "", "", (Int(1000 * sngChrono) / 1000) & " s")
logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCDG.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    FLCDG.Range("A1:K" & UBound(logNom, 1)).CurrentRegion.Borders.LineStyle = xlContinuous
    
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    FLQUA.Range("A2:Q" & derlig).VALUE = quaPos
    'MsgBox UBound(Res, 1) & ":" & UBound(Res, 2) & "<>" & UBound(cInteriorColor, 1) & ":" & UBound(cInteriorColor, 2)
    'FLCIB.Range("A1:" & DecAlph(UBound(Res, 2)) & UBound(Res, 1)).Interior.Color = cInteriorColor
    '''FLCIB.Cells.Clear
    '''FLCIB.Range("A1:" & DecAlph(UBound(Res, 2)) & UBound(Res, 1)).Value = Res
    resColoriage = coloriage(1, FLICI, FLCDG)
    ' Détermination de la ligne de la feuille en cours de traitement
    l = getLineFrom(sheet2process, FLCTL, g_CONTROL_DATA_L, g_CONTROL_DATA_C)
    FLCTL.Cells(l, g_CONTROL_DATA_GEN_C).VALUE = "" & nbliginit & ";" & nbcolinit
    FLCTL.Cells(l, g_CONTROL_DATA_GEN_C + 1).VALUE = "" & nbliggener & ";" & nbcolgener
    FLCTL.Cells(l, g_CONTROL_DATA_GEN_C + 2).VALUE = "" & resColoriage
    FLCTL.Cells(l, g_CONTROL_DATA_GEN_C + 3).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
    'If FLCTL.Cells(9, 5).VALUE = "" Then
        'FLCTL.Cells(9, 5).VALUE = "" & nbliginit & ";" & nbcolinit
        'FLCTL.Cells(9, 6).VALUE = "" & nbliggener & ";" & nbcolgener
        'FLCTL.Cells(9, 7).VALUE = "" & resColoriage
        'FLCTL.Cells(9, 8).VALUE = "" & Round((Int(1000 * sngChrono) / 1000))
    'Else
        'FLCTL.Cells(9, 5).VALUE = FLCTL.Cells(9, 5).VALUE & Chr(10) & nbliginit & ";" & nbcolinit
        'FLCTL.Cells(9, 6).VALUE = FLCTL.Cells(9, 6).VALUE & Chr(10) & nbliggener & ";" & nbcolgener
        'FLCTL.Cells(9, 7).VALUE = FLCTL.Cells(9, 7).VALUE & Chr(10) & resColoriage
        'FLCTL.Cells(9, 8).VALUE = FLCTL.Cells(9, 8).VALUE & Chr(10) & Round((Int(1000 * sngChrono) / 1000))
    'End If
    '''''FLCTL.Select
'''''Application.StatusBar = False
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "setSheet"): Exit Sub
errorHandlerWsNotExists: Call onErrDo("La feuille n'existe pas", "setSheet"): Exit Sub
End Sub
Function isAquantity(qua As String, list() As String) As String()
' Retourne la liste des quantités de qua de list pour tous les domaines TIME ou qua est défini
    Dim res() As String
    ReDim res(0)
    quat = Split(qua, "(")(0)
    'MsgBox quat & Chr(10) & Join(list, Chr(10))
    For I = LBound(list) To UBound(list)
        If list(I) = quat Then
            If UBound(res) = 0 Then
                ReDim res(1 To 1)
            Else
                ReDim Preserve res(1 To UBound(res) + 1)
            End If
            res(UBound(res)) = I & "@" & list(I)
        End If
    Next
    isAquantity = res
End Function
Function isAquantity2(qua As String, list() As String) As String()
' Retourne la liste des quantités de qua de list pour tous les domaines TIME ou qua est défini
    Dim res() As String
    ReDim res(0)
    quat = Split(qua, "(")(0)
    'MsgBox quat & Chr(10) & Join(list, Chr(10))
    isAquantity2 = Filter(list, quat, True)
    'If UBound(resf) > -1 Then
        'ReDim res(1 To (UBound(resf) + 1))
        'For i = LBound(resf) To UBound(resf)
            'res(i + 1) = resf(i)
        'Next
    'End If
    'isAquantity2 = res
End Function


'
Function iterColOrLigOld(claNoligNocol As String, NoColOrLig As Integer, timeslu() As String, NOMENCLALU() As String, areaL() As String, scenarioL() As String, FLCIBLU As Worksheet, colOrLig As String) As Variant
    Dim resultat(0 To 1) As Variant
    Dim errorList(1 To 1) As String
    errorList(1) = "ERROR"
    'Dim rienList(1 To 1) As String
    'rienList(1) = "RIEN"
    'resultat(0) = rienList
    Dim perimeters() As String
    Dim perimetersN() As String
    Dim itera As String
    Dim iteraSansN As String
    Dim numlig As Integer
    Dim lig2process As String
    Dim listEach1() As Variant
    Dim listEach12(1 To 2) As Variant
    Dim listEach23(1 To 2) As Variant
    Dim listEach2() As Variant
    Dim listEach3() As Variant
    ReDim listEach1(0)
    Dim each1 As String
    Dim each2 As String
    Dim each3 As String
    Dim neach1 As String
    Dim neach2 As String
    Dim neach3 As String
    Dim listLines() As String
    ReDim listLines(0)
    Dim listExtended() As String
    ReDim listExtended(1 To 1)
    listExtended(1) = "DEBUT"
    ' détection des itérations
    If UBound(Split(claNoligNocol, "([")) < 1 Then
        ' ERROR
        resultat(0) = errorList
        iterColOrLig = resultat
        Exit Function
    Else
        itera = Split(Split(claNoligNocol, "([")(1), "]=")(0)
    End If
    itera = Replace(itera, "][", "].[")
    iteraSansN = itera
'If colOrLig = "L" Then MsgBox itera & Chr(10) & claNoligNocol & Chr(10) & NoColOrLig
    For I = 1 To 9
        iteraSansN = Replace(iteraSansN, "EACH" & I, "EACH")
        iteraSansN = Replace(iteraSansN, "LEAF" & I, "LEAF")
        iteraSansN = Replace(iteraSansN, "DESC" & I, "DESC")
    Next
    numlig = NoColOrLig
    ' Pour que cela fonctionne avec ][ et ].[
    iteraSansN = Replace(iteraSansN, "][", "].[")
    splIteraSansNDim = Split(iteraSansN, "].[")
    splIteraDim = Split(itera, "].[")
    Dim iteraDim As String
    Dim iteraNDim As String
    Dim iteraNDimFirst As String
    Dim listD() As String
    Dim listIter() As Integer
    Dim un As Integer
    Dim deux As Integer
    Dim trois As Integer
'If colOrLig = "L" Then MsgBox iteraSansN & Chr(10) & claNoligNocol & Chr(10) & NoColOrLig & Chr(10) & LBound(splIteraSansNDim) & ":" & UBound(splIteraSansNDim)
    For spd = LBound(splIteraSansNDim) To UBound(splIteraSansNDim)
        iteraDim = splIteraDim(spd)
        iteraNDim = splIteraSansNDim(spd)
'If colOrLig = "L" Then MsgBox ">>>" & iteraDim & Chr(10) & iteraNDim
        If iteraDim Like "*$*" Then
            ' cas du temps
            iteraNDimFirst = Split(iteraNDim, ">")(0)
            perimeters = getTimes(iteraNDim, timeslu)
'MsgBox iteraNDim & Chr(10) & Chr(10) & Join(timesLu, Chr(10))
        Else
            listD = NOMENCLALU
            If Left(iteraNDim, 1) = "a" Then listD = areaL
            If Left(iteraNDim, 1) = "s" Then listD = scenarioL
            perimeters = getExtended("feu", True, iteraNDim, listD, numlig, "QUANTITE", FLCIBLU.NAME, colOrLig)
'If colOrLig = "C" Then
'MsgBox iteraNDim & Chr(10) & Join(perimeters, Chr(10))
        End If
        If UBound(perimeters) = 0 Then
            resultat(0) = errorList
            iterColOrLig = resultat
            Exit Function
        End If
        ReDim perimetersN(LBound(perimeters) To UBound(perimeters))
        splItera = Split(iteraDim, ".")
        perencours = ""
        facencours = ""
        indEach = ""
        indEach123 = ""
        txtEACH = ""
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10)) & Chr(10) & Chr(10) & Join(splItera, Chr(10))
        For pe = LBound(perimeters) To UBound(perimeters)
            splP = Split(perimeters(pe), ".")
            perencours = ""
            For si = LBound(splItera) To UBound(splItera)
                facencours = splP(si)
                For eacht = 1 To 3
                    If eacht = 1 Then txtEACH = "EACH"
                    If eacht = 2 Then txtEACH = "LEAF"
                    If eacht = 3 Then txtEACH = "DESC"
                    For eachn = 1 To 4
                        If eachn = 4 Then
                            indEach = ""
                            indEach123 = "1"
                        Else
                            indEach = "" & eachn
                            indEach123 = "" & eachn
                        End If
                        If splItera(si) Like "*" & txtEACH & indEach & "*" Then
                            splPerim = Split(iteraDim, ".")
                            avantEACH = Left(splItera(si), InStr(splItera(si), txtEACH & indEach) - 1)
                            facencours = avantEACH & "THIS" & indEach123 & "=" & Right(splP(si), Len(splP(si)) - Len(avantEACH))
                            Exit For
                        End If
                    Next
                Next
                If si = LBound(splItera) Then
                    perencours = facencours
                Else
                    perencours = perencours & "." & facencours
                End If
            Next
            perimetersN(pe) = perencours
        Next
        perimeters = perimetersN
        If spd = LBound(splIteraSansNDim) Then perimeters1 = perimeters
        If spd = UBound(splIteraSansNDim) Then perimeters2 = perimeters
    Next
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10))
    If UBound(splIteraSansNDim) > 0 Then
        ' Reconstitution d'un produit cartésien
        ' a compléter pour les cas ][
        ReDim perimeters(1 To UBound(perimeters1) * UBound(perimeters2))
        Dim ss As Integer
        ss = 0
        For s1 = LBound(perimeters1) To UBound(perimeters1)
            For s2 = LBound(perimeters2) To UBound(perimeters2)
                ss = ss + 1
                perimeters(ss) = perimeters1(s1) & "." & perimeters2(s2)
            Next
        Next
    End If
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10))
    ' préparation des itérations
    each1 = ""
    each2 = ""
    ' détermination des facteurs d'itération
    splPeri = Split(perimeters(LBound(perimeters)), ".")
    ReDim listIter(LBound(splPeri) To UBound(splPeri))
'If colOrLig = "L" Then MsgBox LBound(listIter) & ":" & UBound(listIter)
    un = -1
    deux = -1
    trois = -1
    Dim listIter123() As Integer
    ReDim listIter123(1 To 3)
    For s = LBound(splPeri) To UBound(splPeri)
        listIter(s) = -1
        If splPeri(s) Like "*THIS1*" Then
            listIter(s) = 1
            un = s
            listIter123(1) = s + 1
        End If
        If splPeri(s) Like "*THIS2*" Then
            listIter(s) = 2
            deux = s
            listIter123(2) = s + 1
        End If
        If splPeri(s) Like "*THIS3*" Then
            listIter(s) = 3
            trois = s
            listIter123(3) = s + 1
        End If
    Next
'If colOrLig = "L" Then MsgBox UBound(listIter) & Chr(10) & un & Chr(10) & deux & Chr(10) & trois & Chr(10) & Join(perimeters, Chr(10))
    
    Dim le12(1 To 2) As Variant
    Dim le23(1 To 2) As Variant
    Dim le34(1 To 2) As Variant
    Dim le1() As Variant
    Dim le2() As Variant
    Dim le3() As Variant
    Dim le4() As Variant
    ReDim le1(0)
    ReDim le2(0)
    ReDim le3(0)
    ReDim le4(0)
    Dim pos1 As Integer
    pos1 = 0
    Dim pos2 As Integer
    pos2 = 0
    Dim pos3 As Integer
    pos3 = 0
    For p1 = LBound(perimeters) To UBound(perimeters)
        splPeri = Split(perimeters(p1), ".")
        If un > -1 Then
            neach1 = splPeri(un)
            ReDim le2(0)
            ' recherche position de neach1
            If UBound(le1) = 0 Then
                ReDim le1(1 To 1)
                le12(1) = perimeters(p1)
                le12(2) = le2
                le1(UBound(le1)) = le12
                pos1 = UBound(le1)
            Else
                pos1 = 0
                For l = LBound(le1) To UBound(le1)
                    per = le1(l)(1)
                    splper = Split(per, ".")
                    If splper(un) = neach1 Then
                        pos1 = l
                        Exit For
                    End If
                Next
                If pos1 = 0 Then
                    ReDim Preserve le1(1 To UBound(le1) + 1)
                    le12(1) = perimeters(p1)
                    le12(2) = le2
                    le1(UBound(le1)) = le12
                    pos1 = UBound(le1)
                End If
            End If
            If deux > -1 Then
                neach2 = splPeri(deux)
                ReDim le3(0)
                If UBound(le1(pos1)(2)) = 0 Then
                    ReDim le2(1 To 1)
                    le23(1) = perimeters(p1)
                    le23(2) = le3
                    le2(UBound(le2)) = le23
                    le1(pos1)(2) = le2
                    pos2 = UBound(le2)
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & "nouveau" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                Else
                    pos2 = 0
                    For l = LBound(le1(pos1)(2)) To UBound(le1(pos1)(2))
                        per = le1(pos1)(2)(l)(1)
                        splper = Split(per, ".")
                        If splper(deux) = neach2 Then
                            pos2 = l
                            Exit For
                        End If
                    Next
                    If pos2 = 0 Then
                        le2 = le1(pos1)(2)
                        ReDim Preserve le2(1 To UBound(le2) + 1)
                        le23(1) = perimeters(p1)
                        le23(2) = le3
                        le2(UBound(le2)) = le23
                        le1(pos1)(2) = le2
                        pos2 = UBound(le2)
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & splper(deux) & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                    End If
                End If
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & "resul" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                If trois > -1 Then
                    neach3 = splPeri(trois)
                    ReDim le4(0)
'If colOrLig = "L" Then MsgBox neach1 & ":" & neach2 & ":" & neach3 & Chr(10) & UBound(le1(pos1)(2)) & Chr(10) & UBound(le1(pos1)(2)(pos2)(2))
                    If UBound(le1(pos1)(2)(pos2)(2)) = 0 Then
                        ReDim le3(1 To 1)
                        le34(1) = perimeters(p1)
                        le34(2) = le4
                        le3(UBound(le3)) = le34
                        le1(pos1)(2)(pos2)(2) = le3
                        pos3 = UBound(le3)
'If colOrLig = "L" Then MsgBox "nouveau" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                    Else
                        pos3 = 0
                        For l = LBound(le1(pos1)(2)(pos2)(2)) To UBound(le1(pos1)(2)(pos2)(2))
                            per = le1(pos1)(2)(pos2)(2)(l)(1)
                            splper = Split(per, ".")
                            If splper(trois) = neach3 Then
                                pos3 = l
                                Exit For
                            End If
                        Next
                        If pos2 = 0 Then
                            le3 = le1(pos1)(2)(pos2)(2)
                            ReDim Preserve le3(1 To UBound(le3) + 1)
                            le34(1) = perimeters(p1)
                            le34(2) = le4
                            le3(UBound(le3)) = le34
                            le1(pos1)(2)(pos2)(2) = le3
                            pos3 = UBound(le3)
'If colOrLig = "L" Then MsgBox "suite" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                        End If
                    End If
                
                End If
            End If
        End If
    Next

'If colOrLig = "L" Then
'For i = LBound(le1) To UBound(le1)
    'For j = LBound(le1(i)(2)) To UBound(le1(i)(2))
'''MsgBox i & ":" & j & Chr(10) & le1(i)(2)(j)(1)
    'Next
'Next
'End If

    Dim sis As Integer
    each1 = ""
    each2 = ""
'If un = -1 Then MsgBox Join(perimeters, Chr(10))
    For p = LBound(perimeters) To UBound(perimeters)
        splPeri = Split(perimeters(p), ".")
        If un > -1 Then
        neach1 = splPeri(un)
        If neach1 <> each1 Then
            ' Traitement du niveau 1 de l'itération
            each2 = ""
            ReDim listEach2(0)
            If UBound(listEach1) > 0 Then
                ReDim Preserve listEach1(1 To UBound(listEach1) + 1)
            Else
                ReDim listEach1(1 To 1)
            End If
            listEach12(1) = perimeters(p)
            listEach1(UBound(listEach1)) = listEach12
'If colOrLig = "L" Then MsgBox "n1 " & UBound(listEach1) & ":" & listEach1(UBound(listEach1))(1)
        End If
        If deux > -1 Then
            neach2 = splPeri(deux)
            If neach2 <> each2 Then
                each3 = ""
                ReDim listEach3(0)
                If UBound(listEach2) > 0 Then
                    ReDim Preserve listEach2(1 To UBound(listEach2) + 1)
                Else
                    ReDim listEach2(1 To 1)
                End If
                listEach23(1) = perimeters(p)
                listEach2(UBound(listEach2)) = listEach23
                listEach12(2) = listEach2
                listEach1(UBound(listEach1)) = listEach12
'If colOrLig = "L" Then MsgBox "n2 " & UBound(listEach1) & ":" & UBound(listEach2) & "::" & listEach1(UBound(listEach1))(2)(UBound(listEach2))(1)
            End If
        Else
            neach2 = ""
        End If
        If trois > -1 Then
            each3 = splPeri(trois)
        End If
        each1 = neach1
        each2 = neach2
        End If
    Next
'If colOrLig = "L" Then
'MsgBox LBound(listEach1) & ":" & UBound(listEach1)
    'For ii = LBound(listEach1) To UBound(listEach1)
'MsgBox "1er niveau=" & ii & ":" & listEach1(ii)(1)
'& Chr(10) & LBound(listEach1(ii)(2)) & ":" & UBound(listEach1(ii)(2))
        'For jj = LBound(listEach1(ii)(2)) To UBound(listEach1(ii)(2))
            'If colOrLig = "L" Then
            'MsgBox "2ème niveau=" & ii & ":" & jj & ":" & listEach1(ii)(2)(jj)(1)
        'Next
    'Next
'End If
    resultat(0) = listExtended
    '''resultat(1) = listEach1
    resultat(1) = le1
    iterColOrLig = resultat
End Function

Function iterColOrLig(claNoligNocol As String, NoColOrLig As Integer, timeslu() As String, NOMENCLALU() As String, areaL() As String, scenarioL() As String, FLCIBLU As Worksheet, colOrLig As String) As Variant
    'Arguments      :
    'claNoligNocol  : expression de l'itération du type [A>EACH1][B>EACH2]=...
    'NoColOrLig     : numéro de colonne ou ligne de la feuille générique
    'timeslu        :
    'NOMENCLALU     :
    'areaL          :
    'scenarioL      :
    'FLCIBLU        : feuille générique
    'colOrLig       : L en boucle ligne, C en boucle colonne
    'Sortie         :
    '               :
    Dim resultat(0 To 1) As Variant
    Dim errorList(1 To 1) As String
    Dim listPerimeters() As Variant
    errorList(1) = "ERROR"
    'Dim rienList(1 To 1) As String
    'rienList(1) = "RIEN"
    'resultat(0) = rienList
    Dim perimeters() As String
    Dim perimetersN() As String
    Dim itera As String
    Dim iteraSansN As String
    Dim numlig As Integer
    Dim lig2process As String
    Dim listEach1() As Variant
    Dim listEach12(1 To 2) As Variant
    Dim listEach23(1 To 2) As Variant
    Dim listEach2() As Variant
    Dim listEach3() As Variant
    ReDim listEach1(0)
    Dim each1 As String
    Dim each2 As String
    Dim each3 As String
    Dim neach1 As String
    Dim neach2 As String
    Dim neach3 As String
    Dim listLines() As String
    ReDim listLines(0)
    Dim listExtended() As String
    ReDim listExtended(1 To 1)
    listExtended(1) = "DEBUT"
    ' détection des itérations
    If UBound(Split(claNoligNocol, "([")) < 1 Then
        ' ERROR
        resultat(0) = errorList
        iterColOrLig = resultat
        Exit Function
    Else
        itera = Split(Split(claNoligNocol, "([")(1), "]=")(0)
    End If
    itera = Replace(itera, "][", "].[")
    iteraSansN = itera
'If colOrLig = "L" Then MsgBox itera & Chr(10) & claNoligNocol & Chr(10) & NoColOrLig
    For I = 1 To 9
        iteraSansN = Replace(iteraSansN, "EACH" & I, "EACH")
        iteraSansN = Replace(iteraSansN, "LEAF" & I, "LEAF")
        iteraSansN = Replace(iteraSansN, "DESC" & I, "DESC")
    Next
    numlig = NoColOrLig
    ' Pour que cela fonctionne avec ][ et ].[
    iteraSansN = Replace(iteraSansN, "][", "].[")
    splIteraSansNDim = Split(iteraSansN, "].[")
    splIteraDim = Split(itera, "].[")
    Dim iteraDim As String
    Dim iteraNDim As String
    Dim iteraNDimFirst As String
    Dim listD() As String
    Dim listIter() As Integer
    Dim un As Integer
    Dim deux As Integer
    Dim trois As Integer
'If colOrLig = "C" Then MsgBox iteraSansN & Chr(10) & claNoligNocol & Chr(10) & NoColOrLig & Chr(10) & LBound(splIteraSansNDim) & ":" & UBound(splIteraSansNDim)
    ReDim listPerimeters(LBound(splIteraSansNDim) To UBound(splIteraSansNDim))
    For spd = LBound(splIteraSansNDim) To UBound(splIteraSansNDim)
        iteraDim = splIteraDim(spd)
        iteraNDim = splIteraSansNDim(spd)
'If colOrLig = "L" Then MsgBox ">>>" & iteraDim & Chr(10) & iteraNDim
        If iteraDim Like "*$*" Then
            ' cas du temps
            iteraNDimFirst = Split(iteraNDim, ">")(0)
            perimeters = getTimes(iteraNDim, timeslu)
'MsgBox iteraNDim & Chr(10) & Chr(10) & Join(timesLu, Chr(10))
        Else
            listD = NOMENCLALU
            If Left(iteraNDim, 1) = "a" Then listD = areaL
            If Left(iteraNDim, 1) = "s" Then listD = scenarioL
            perimeters = getExtended("feu", True, iteraNDim, listD, numlig, "QUANTITE", FLCIBLU.NAME, colOrLig)
'If colOrLig = "C" Then MsgBox Join(perimeters, Chr(10))
'If colOrLig = "C" Then
'MsgBox iteraNDim & Chr(10) & Join(perimeters, Chr(10))
        End If
        If UBound(perimeters) = 0 Then
            resultat(0) = errorList
            iterColOrLig = resultat
            Exit Function
        End If
        ReDim perimetersN(LBound(perimeters) To UBound(perimeters))
        splItera = Split(iteraDim, ".")
        perencours = ""
        facencours = ""
        indEach = ""
        indEach123 = ""
        txtEACH = ""
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10)) & Chr(10) & Chr(10) & Join(splItera, Chr(10))
        For pe = LBound(perimeters) To UBound(perimeters)
            splP = Split(perimeters(pe), ".")
            perencours = ""
            For si = LBound(splItera) To UBound(splItera)
                facencours = splP(si)
                For eacht = 1 To 3
                    If eacht = 1 Then txtEACH = "EACH"
                    If eacht = 2 Then txtEACH = "LEAF"
                    If eacht = 3 Then txtEACH = "DESC"
                    For eachn = 1 To 4
                        If eachn = 4 Then
                            indEach = ""
                            indEach123 = "1"
                        Else
                            indEach = "" & eachn
                            indEach123 = "" & eachn
                        End If
                        If splItera(si) Like "*" & txtEACH & indEach & "*" Then
                            splPerim = Split(iteraDim, ".")
                            avantEACH = Left(splItera(si), InStr(splItera(si), txtEACH & indEach) - 1)
                            facencours = avantEACH & "THIS" & indEach123 & "=" & Right(splP(si), Len(splP(si)) - Len(avantEACH))
                            Exit For
                        End If
                    Next
                Next
                If si = LBound(splItera) Then
                    perencours = facencours
                Else
                    perencours = perencours & "." & facencours
                End If
            Next
            perimetersN(pe) = perencours
        Next
        perimeters = perimetersN
        If spd = LBound(splIteraSansNDim) Then perimeters1 = perimeters
        If spd = UBound(splIteraSansNDim) Then perimeters2 = perimeters
        listPerimeters(spd) = perimeters
    Next
'If colOrLig = "C" Then MsgBox "after1" & Chr(10) & UBound(splIteraSansNDim) & ":" & UBound(splIteraSansNDim)
'If colOrLig = "C" Then MsgBox "1after1" & Chr(10) & Join(perimeters1, Chr(10))
'If colOrLig = "C" Then MsgBox "2after1" & Chr(10) & Join(perimeters2, Chr(10))
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10))
    Dim ss As Integer
    If UBound(splIteraSansNDim) = 1 Then
'If colOrLig = "C" Then MsgBox (UBound(listPerimeters(0)) * UBound(listPerimeters(1)))
        ' Reconstitution d'un produit cartésien
        ' a compléter pour les cas ][
        ReDim perimeters(1 To UBound(listPerimeters(0)) * UBound(listPerimeters(1)))
        ss = 0
        For s1 = LBound(listPerimeters(0)) To UBound(listPerimeters(0))
            For s2 = LBound(listPerimeters(1)) To UBound(listPerimeters(1))
                ss = ss + 1
                perimeters(ss) = listPerimeters(0)(s1) & "." & listPerimeters(1)(s2)
            Next
        Next
    End If
'If colOrLig = "C" Then MsgBox "ici"
    If UBound(splIteraSansNDim) = 2 Then
        ' Reconstitution d'un produit cartésien
        ' a compléter pour les cas ][
        ReDim perimeters(1 To UBound(listPerimeters(0)) * UBound(listPerimeters(1)) * UBound(listPerimeters(2)))
        ss = 0
        For s1 = LBound(listPerimeters(0)) To UBound(listPerimeters(0))
            For s2 = LBound(listPerimeters(1)) To UBound(listPerimeters(1))
                For s3 = LBound(listPerimeters(2)) To UBound(listPerimeters(2))
                    ss = ss + 1
                    perimeters(ss) = listPerimeters(0)(s1) & "." & listPerimeters(1)(s2) & "." & listPerimeters(2)(s3)
                Next
            Next
        Next
    End If
'If colOrLig = "C" Then MsgBox "iterColOrLig" & Chr(10) & Join(perimeters, Chr(10))
'If colOrLig = "L" Then MsgBox Join(perimeters, Chr(10))
    ' préparation des itérations
    each1 = ""
    each2 = ""
    ' détermination des facteurs d'itération
    splPeri = Split(perimeters(LBound(perimeters)), ".")
    ReDim listIter(LBound(splPeri) To UBound(splPeri))
'If colOrLig = "L" Then MsgBox LBound(listIter) & ":" & UBound(listIter)
    un = -1
    deux = -1
    trois = -1
    Dim listIter123() As Integer
    ReDim listIter123(1 To 3)
    For s = LBound(splPeri) To UBound(splPeri)
        listIter(s) = -1
        If splPeri(s) Like "*THIS1*" Then
            listIter(s) = 1
            un = s
            listIter123(1) = s + 1
        End If
        If splPeri(s) Like "*THIS2*" Then
            listIter(s) = 2
            deux = s
            listIter123(2) = s + 1
        End If
        If splPeri(s) Like "*THIS3*" Then
            listIter(s) = 3
            trois = s
            listIter123(3) = s + 1
        End If
    Next
'If colOrLig = "L" Then MsgBox UBound(listIter) & Chr(10) & un & Chr(10) & deux & Chr(10) & trois & Chr(10) & Join(perimeters, Chr(10))
    
    Dim le12(1 To 2) As Variant
    Dim le23(1 To 2) As Variant
    Dim le34(1 To 2) As Variant
    Dim le1() As Variant
    Dim le2() As Variant
    Dim le3() As Variant
    Dim le4() As Variant
    ReDim le1(0)
    ReDim le2(0)
    ReDim le3(0)
    ReDim le4(0)
    Dim pos1 As Integer
    pos1 = 0
    Dim pos2 As Integer
    pos2 = 0
    Dim pos3 As Integer
    pos3 = 0
'MsgBox un & Chr(10) & deux & Chr(10) & trois & Chr(10) & Join(perimeters, Chr(10))
    For p1 = LBound(perimeters) To UBound(perimeters)
        splPeri = Split(perimeters(p1), ".")
        If un > -1 Then
            neach1 = splPeri(un)
            ReDim le2(0)
            ' recherche position de neach1
            If UBound(le1) = 0 Then
                ReDim le1(1 To 1)
                le12(1) = perimeters(p1)
                le12(2) = le2
                le1(UBound(le1)) = le12
                pos1 = UBound(le1)
            Else
                pos1 = 0
                For l = LBound(le1) To UBound(le1)
                    per = le1(l)(1)
                    splper = Split(per, ".")
                    If splper(un) = neach1 Then
                        pos1 = l
                        Exit For
                    End If
                Next
                If pos1 = 0 Then
                    ReDim Preserve le1(1 To UBound(le1) + 1)
                    le12(1) = perimeters(p1)
                    le12(2) = le2
                    le1(UBound(le1)) = le12
                    pos1 = UBound(le1)
                End If
            End If
            If deux > -1 Then
                neach2 = splPeri(deux)
                ReDim le3(0)
                If UBound(le1(pos1)(2)) = 0 Then
                    ReDim le2(1 To 1)
                    le23(1) = perimeters(p1)
                    le23(2) = le3
                    le2(UBound(le2)) = le23
                    le1(pos1)(2) = le2
                    pos2 = UBound(le2)
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & "nouveau" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                Else
                    pos2 = 0
                    For l = LBound(le1(pos1)(2)) To UBound(le1(pos1)(2))
                        per = le1(pos1)(2)(l)(1)
                        splper = Split(per, ".")
                        If splper(deux) = neach2 Then
                            pos2 = l
                            Exit For
                        End If
                    Next
                    If pos2 = 0 Then
                        le2 = le1(pos1)(2)
                        ReDim Preserve le2(1 To UBound(le2) + 1)
                        le23(1) = perimeters(p1)
                        le23(2) = le3
                        le2(UBound(le2)) = le23
                        le1(pos1)(2) = le2
                        pos2 = UBound(le2)
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & splper(deux) & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                    End If
                End If
'If colOrLig = "L" Then MsgBox perimeters(p1) & ":" & "resul" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                If trois > -1 Then
                    neach3 = splPeri(trois)
                    ReDim le4(0)
'If colOrLig = "L" Then MsgBox neach1 & ":" & neach2 & ":" & neach3 & Chr(10) & UBound(le1(pos1)(2)) & Chr(10) & UBound(le1(pos1)(2)(pos2)(2))
                    If UBound(le1(pos1)(2)(pos2)(2)) = 0 Then
                        ReDim le3(1 To 1)
                        le34(1) = perimeters(p1)
                        le34(2) = le4
                        le3(UBound(le3)) = le34
                        le1(pos1)(2)(pos2)(2) = le3
                        pos3 = UBound(le3)
'If colOrLig = "L" Then MsgBox "nouveau" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                    Else
                        pos3 = 0
                        For l = LBound(le1(pos1)(2)(pos2)(2)) To UBound(le1(pos1)(2)(pos2)(2))
                            per = le1(pos1)(2)(pos2)(2)(l)(1)
                            splper = Split(per, ".")
                            If splper(trois) = neach3 Then
                                pos3 = l
                                Exit For
                            End If
                        Next
                        If pos3 = 0 Then
                            le3 = le1(pos1)(2)(pos2)(2)
                            ReDim Preserve le3(1 To UBound(le3) + 1)
                            le34(1) = perimeters(p1)
                            le34(2) = le4
                            le3(UBound(le3)) = le34
                            le1(pos1)(2)(pos2)(2) = le3
                            pos3 = UBound(le3)
'If colOrLig = "L" Then MsgBox "suite" & Chr(10) & pos1 & ":" & pos2 & Chr(10) & perimeters(p1) & Chr(10) & neach1 & Chr(10) & neach2
                        End If
                    End If
                
                End If
            End If
        End If
    Next

'If colOrLig = "C" Then
'For i = LBound(le1) To UBound(le1)
    'For j = LBound(le1(i)(2)) To UBound(le1(i)(2))
'MsgBox i & ":" & j & Chr(10) & le1(i)(2)(j)(1)
    'Next
'Next
'End If

    Dim sis As Integer
    each1 = ""
    each2 = ""
'If un = -1 Then MsgBox Join(perimeters, Chr(10))
    For p = LBound(perimeters) To UBound(perimeters)
        splPeri = Split(perimeters(p), ".")
        If un > -1 Then
        neach1 = splPeri(un)
        If neach1 <> each1 Then
            ' Traitement du niveau 1 de l'itération
            each2 = ""
            ReDim listEach2(0)
            If UBound(listEach1) > 0 Then
                ReDim Preserve listEach1(1 To UBound(listEach1) + 1)
            Else
                ReDim listEach1(1 To 1)
            End If
            listEach12(1) = perimeters(p)
            listEach1(UBound(listEach1)) = listEach12
'If colOrLig = "L" Then MsgBox "n1 " & UBound(listEach1) & ":" & listEach1(UBound(listEach1))(1)
        End If
        If deux > -1 Then
            neach2 = splPeri(deux)
            If neach2 <> each2 Then
                each3 = ""
                ReDim listEach3(0)
                If UBound(listEach2) > 0 Then
                    ReDim Preserve listEach2(1 To UBound(listEach2) + 1)
                Else
                    ReDim listEach2(1 To 1)
                End If
                listEach23(1) = perimeters(p)
                listEach2(UBound(listEach2)) = listEach23
                listEach12(2) = listEach2
                listEach1(UBound(listEach1)) = listEach12
'If colOrLig = "L" Then MsgBox "n2 " & UBound(listEach1) & ":" & UBound(listEach2) & "::" & listEach1(UBound(listEach1))(2)(UBound(listEach2))(1)
            End If
        Else
            neach2 = ""
        End If
        If trois > -1 Then
            each3 = splPeri(trois)
        End If
        each1 = neach1
        each2 = neach2
        End If
    Next
'If colOrLig = "C" Then
'MsgBox LBound(listEach1) & ":" & UBound(listEach1)
    'For ii = LBound(listEach1) To UBound(listEach1)
'MsgBox "1er niveau=" & ii & ":" & listEach1(ii)(1) & Chr(10) & LBound(listEach1(ii)) & ":" & UBound(listEach1(ii))
        'For jj = LBound(listEach1(ii)(2)) To UBound(listEach1(ii)(2))
            'If colOrLig = "L" Then
            'MsgBox "2ème niveau=" & ii & ":" & jj & ":" & listEach1(ii)(2)(jj)(1)
        'Next
    'Next
'End If
    resultat(0) = listExtended
    '''resultat(1) = listEach1
'If colOrLig = "C" Then MsgBox Join(resultat(0), Chr(10))
'If colOrLig = "C" Then MsgBox LBound(le1(1)(2)) & ":" & UBound(le1(1)(2))
'If colOrLig = "C" Then MsgBox ">" & le1(1)(1) & "<"
    resultat(1) = le1
'0:1                                    LBound(resultat) & ":" & UBound(resultat)
'DEBUT                                  Join(resultat(0), Chr(10))
'1:3                                    LBound(resultat(1)) & ":" & UBound(resultat(1))
 '1:2                                   LBound(resultat(1)(1)) & ":" & UBound(resultat(1)(1))
 '1:2                                   LBound(resultat(1)(2)) & ":" & UBound(resultat(1)(2))
 '1:2                                   LBound(resultat(1)(3)) & ":" & UBound(resultat(1)(3))
 'A>THIS1=a1.B>THIS2=b1.C>THIS3=c1      resultat(1)(1)(1)
 '1:2                                   LBound(resultat(1)(1)(2)) & ":" & UBound(resultat(1)(1)(2))
 '1:2                                   LBound(resultat(1)(1)(2)(1)) & ":" & UBound(resultat(1)(1)(2)(1))
''MsgBox LBound(resultat(1)(1)(2)) & ":" & UBound(resultat(1)(1)(2))
''MsgBox resultat(1)(1)(2)(1)(1)
'MsgBox claNoligNocol & ":" & colOrLig
'MsgBox LBound(resultat(1)) & ":" & UBound(resultat(1))
'MsgBox "0=" & Join(resultat(0), Chr(10)) & Chr(10) & "11=" & UBound(resultat(1)(1))
    iterColOrLig = resultat
End Function
Function getMatch(listExtend() As String, listE1() As Variant, listLin() As String, num As Integer, nomList() As String, areList() As String, sceList() As String) As String()
    On Error GoTo errorHandler
    listExtend = listLin
'MsgBox Join(listLin, Chr(10)) & Chr(10) & Chr(10) & Join(listExtend, Chr(10))
    Dim listExtendEnd() As String
    Dim nle() As String
    ReDim listExtendEnd(0)
    Dim newProduit As String
    newProduit = ""
    Dim ajout As String
    Dim list() As String
'If num = 7 Then MsgBox LBound(listE1) & ":" & UBound(listE1) & "=Début=" & Chr(10) & Join(listExtend, Chr(10))
    For I = LBound(listE1) To UBound(listE1)
        listExtend = listLin
'MsgBox Join(listExtend, Chr(10)) & Chr(10) & "___" & Chr(10) & listE1(I)(1)
        For l = LBound(listExtend) To UBound(listExtend)
            If listExtend(l) Like "*.1*" Then
                splE1 = Split(listExtend(l), "[")(1)
                splE2 = Split(splE1, "]")(0)
                If splE2 = "" Then
                    listExtend(l) = l & ".1-" & I & "+" & listE1(I)(1)
                Else
                    listExtend(l) = l & ".1-" & I & "+[" & splE2 & "]" & listE1(I)(1)
                End If
            End If
        Next
''If num = 7 Then MsgBox i & "=1=" & Chr(10) & Join(listExtend, Chr(10))
'MsgBox Join(listExtend, Chr(10))

        If Not IsEmpty(listE1(I)(UBound(listE1(I)))) Then
            If UBound(listE1(I)) > 0 Then
'MsgBox "niveau 2" & "::" & LBound(listE1(I)(2)) & Chr(10) & UBound(listE1(I)(2))
            For j = LBound(listE1(I)(2)) To UBound(listE1(I)(2))
'If num = 7 Then MsgBox i & ":" & j & Chr(10) & listE1(i)(2)(j)(1)
'MsgBox "niveau 2"
                For l = LBound(listExtend) To UBound(listExtend)
                    If listExtend(l) Like "*.2*" Then
                        splE1 = Split(listExtend(l) & "[", "[")(1)
                        If splE1 <> "" Then
                            splE2 = Split(splE1, "]")(0)
                        Else
                            splE2 = ""
                        End If
                        If splE2 = "" Then
                            If j = LBound(listE1(I)(2)) Then
    listExtend(l) = l & ".2-" & I & "_" & LBound(listE1(I)(2)) & "+" & listE1(I)(2)(j)(1) & "|"
                            Else
    listExtend(l) = l & ".2-" & I & "_" & LBound(listE1(I)(2)) & "+" & Split(listExtend(l), "+")(1) & "@" & listE1(I)(2)(j)(1) & "|"
                            End If
                        Else
                            If j = LBound(listE1(I)(2)) Then
    listExtend(l) = l & ".2-" & I & "_" & LBound(listE1(I)(2)) & "+[" & splE2 & "]" & listE1(I)(2)(j)(1) & "|"
                            Else
    listExtend(l) = l & ".2-" & I & "_" & LBound(listE1(I)(2)) & "+" & Split(listExtend(l), "+")(1) & "@[" & splE2 & "]" & listE1(I)(2)(j)(1) & "|"
                            End If
                        End If
                    End If
                Next
                ' Traitement du niveau 3
                If UBound(listE1(I)(2)) > 0 Then
                If UBound(listE1(I)(2)(j)) > 0 Then
'MsgBox "niveau 3 " & I & ">" & j & " : " & UBound(listE1(I)(2)(j)) & Chr(10) & Join(listExtend, Chr(10))
''& "::" & LBound(listE1(I)(2)(j)(2)) & Chr(10) & UBound(listE1(I)(2)(j)(2))
                    For k = LBound(listE1(I)(2)(j)(2)) To UBound(listE1(I)(2)(j)(2))
                        For l = LBound(listExtend) To UBound(listExtend)
                            If listExtend(l) Like "*.3*" Then
'MsgBox "listExtend" & Chr(10) & listExtend(l)
'MsgBox I & ">" & j & ">" & k & " niveau 3 : " & listExtend(l)
                                'Dim lastLigne As String
                                'lastLigne = listExtend(l)
                                splE1 = Split(listExtend(l) & "[", "[")(1)
                                If splE1 <> "" Then
                                    splE2 = Split(splE1, "]")(0)
                                Else
                                    splE2 = ""
                                End If
'MsgBox splE1 & Chr(10) & ">" & splE2 & "<k=" & k & ":" & LBound(listE1(I)(2)(j)(2))
                                If splE2 = "" Then
                                    If k = LBound(listE1(I)(2)(j)(2)) Then
'MsgBox listE1(I)(2)(j)(1) & Chr(10) & listE1(I)(2)(j)(2)(k)(1)
    listExtend(l) = l & ".3-" & I & "_" & LBound(listE1(I)(2)(j)(2)) & "+" & listE1(I)(2)(j)(2)(k)(1)
    '''listExtend(l) = l & ".2-" & i & "_" & LBound(listE1(i)(2)) & "+" & listE1(i)(2)(j)(1)
                                    Else
    listExtend(l) = l & ".3-" & I & "_" & LBound(listE1(I)(2)(j)(2)) & "+" & Split(listExtend(l), "+")(1) & "@" & listE1(I)(2)(j)(2)(k)(1)
    '''listExtend(l) = l & ".2-" & i & "_" & LBound(listE1(i)(2)) & "+" & Split(listExtend(l), "+")(1) & "@" & listE1(i)(2)(j)(1)
                                    End If
'MsgBox k & ":VIDE=" & splE1 & ":" & listE1(I)(2)(j)(1) & Chr(10) & Join(listExtend, Chr(10))
                                Else
                                    If k = LBound(listE1(I)(2)(j)(2)) Then
    listExtend(l) = l & ".3-" & I & "_" & LBound(listE1(I)(2)(j)(2)) & "+[" & splE2 & "]" & listE1(I)(2)(j)(2)(k)(1)
    '''listExtend(l) = l & ".2-" & i & "_" & LBound(listE1(i)(2)) & "+[" & splE2 & "]" & listE1(i)(2)(j)(1)
                                    Else
    listExtend(l) = l & ".3-" & I & "_" & LBound(listE1(I)(2)(j)(2)) & "+" & Split(listExtend(l), "+")(1) & "@[" & splE2 & "]" & listE1(I)(2)(j)(2)(k)(1)
    '''listExtend(l) = l & ".2-" & i & "_" & LBound(listE1(i)(2)) & "+" & Split(listExtend(l), "+")(1) & "@[" & splE2 & "]" & listE1(i)(2)(j)(1)
                                    End If
'MsgBox k & ":FULL=" & splE1 & ":" & listE1(I)(2)(j)(1) & Chr(10) & Join(listExtend, Chr(10))
                                End If
                            End If
'If Mid(listExtend(l), 1, 2) = "3." Then MsgBox listE1(I)(2)(j)(2)(k)(1) & Chr(10) & listExtend(l)
                        Next
'MsgBox k & ":::" & Chr(10) & Join(listExtend, Chr(10))
                    Next
                    ' linéarisation sur le niveau 2 juste au dessus
'MsgBox "getMatch0:" & I & j & "/" & UBound(listE1(I)(2)) & Chr(10) & Join(listExtend, Chr(10))
' cette boucle est plus complexe il faut boucler sur les items d'une même ligne
                    For l = LBound(listExtend) To UBound(listExtend)
                        If listExtend(l) Like "*.3*" Then
                            'aaa = listExtend(l - 1)
                            listExtend(l - 1) = Replace(listExtend(l - 1), "||", "|")
                            listExtend(l) = Replace(listExtend(l), "@", "µ")
                            listExtend(l - 1) = Replace(listExtend(l - 1), "|", "µ" & Split(listExtend(l), "+")(1))
'MsgBox I & j & Chr(10) & Chr(10) & listExtend(l - 1) & Chr(10) & Chr(10) & listExtend(l)
                            If j = UBound(listE1(I)(2)) Then listExtend(l) = ""
'MsgBox I & j & Chr(10) & Chr(10) & listExtend(l - 1) & Chr(10) & Chr(10) & listExtend(l)
                        End If
                    Next
                    ReDim nle(0)
                    Dim p As Integer
                    p = 0
'MsgBox "getMatch1:" & I & j & "/" & UBound(listE1(I)(2)) & Chr(10) & Join(listExtend, Chr(10))
'MsgBox "ici" & LBound(listExtend) & ":" & UBound(listExtend)
                    For l = LBound(listExtend) To UBound(listExtend)
                        If listExtend(l) <> "" Then
                            p = p + 1
                            If p = 1 Then
                                ReDim nle(1 To 1)
                            Else
                                ReDim Preserve nle(1 To p)
                            End If
                            nle(p) = listExtend(l) '''Replace(listExtend(l), "|", "")
                        End If
                    Next
'MsgBox "AAA" & Chr(10) & Join(listExtend, Chr(10)) & Chr(10) & "AAA" & Chr(10) & Join(nle, Chr(10)) & Chr(10) & "AAA"
                    listExtend = nle
'MsgBox "FIN" & Chr(10) & Join(listExtend, Chr(10))
                End If
                End If
            Next
            End If
'If num = 7 Then
'MsgBox I & "=2=" & UBound(listE1(I)) & ":" & UBound(listE1(I)(2)) & Chr(10) & Join(listExtend, Chr(10))
        End If
'MsgBox I & "A" & Chr(10) & Join(listExtend, Chr(10))
        ' résolution des THIS
        Dim listExtendN() As String
        ReDim listExtendN(LBound(listExtend) To UBound(listExtend))
'MsgBox Join(listExtend, Chr(10))
'Dim old() As String
'old = listExtend
        For e = LBound(listExtend) To UBound(listExtend)
            listExtend(e) = Replace(listExtend(e), "@", "@1|")
            listExtend(e) = Replace(listExtend(e), "µ", "@2|")
            splAr = Split(listExtend(e), "@")
            newExt = ""
            For se = LBound(splAr) To UBound(splAr)
                aTraiter = splAr(se)
                If aTraiter Like "*]*" Then
                    spllistE = Split(aTraiter, "]")
                    av = spllistE(0)
                    ap = spllistE(1)
                    If av Like "*THIS*" Then
                        nn = ""
                        For x = 1 To 4
                            If x = 4 Then
                                nn = ""
                            Else
                                nn = "" & x
                            End If
                            aptn = Split(ap, "THIS" & nn & "=")
                            If UBound(aptn) > 0 Then
                                tn = Split(Split(aptn(1), ":")(0), ".")(0)
'If num = 45 Then
'MsgBox av & Chr(10) & "THIS" & nn & Chr(10) & tn
                                av = Replace(av, "THIS" & nn, tn)
                            End If
                        Next
                        aTraiter = av & "]" & ap
                    End If
                End If
                If newExt = "" Then
                    newExt = aTraiter
                Else
                    newExt = newExt & "@" & aTraiter
                End If
            Next
            newExt = Replace(newExt, "@2|", "µ")
            newExt = Replace(newExt, "@1|", "@")
            listExtendN(e) = newExt
        Next
'MsgBox Join(old, Chr(10)) & Chr(10) & Chr(10) & Join(listExtendN, Chr(10))
        listExtend = listExtendN
'If num = 45 Then MsgBox Join(listExtend, Chr(10))
        ' Eclatement des @
        lastLin = ""
        lastLev = ""
        Dim listAderouler() As String
        ReDim listAderouler(0)
        Dim newListExend0() As String
        Dim newListExend() As String
        ReDim newListExend0(0)
        ReDim newListExend(0)
        Dim ajoutExt() As String
        Dim ok4add As Boolean
'If currlin = "" Then MsgBox ">>>>>>>>>>>>" & UBound(listExtend) & ">>>" & num & Chr(10) & Join(listExtend, Chr(10))
'MsgBox ">>>>>>>>>>>>" & UBound(listExtend) & ">>>" & num & Chr(10) & Join(listExtend, Chr(10))
        For e = LBound(listExtend) To UBound(listExtend)
            currlin = listExtend(e)
'If currlin = "" Then MsgBox ">>>>>>>>>>>>" & UBound(listExtend) & ">>>" & num & Chr(10) & Join(listLin, Chr(10))
'If num = 45 Then
'If currlin = "" Then
'MsgBox num & "==>" & currlin & Chr(10) & Join(listExtend, Chr(10))
            currSansPlus = Split(Split(currlin, "+")(0), "-")(0)
            currLev = Split(currSansPlus, ".")(1)
            If lastLev = currLev Then
                'on alimente le tableau
                If UBound(listAderouler) = 0 Then
                    ReDim listAderouler(1 To 1)
                Else
                    ReDim Preserve listAderouler(1 To UBound(listAderouler) + 1)
                End If
                listAderouler(UBound(listAderouler)) = lastLin
            Else
                ' cas d'un item différend du précédent donc on alimente les précédents
                If UBound(listAderouler) = 0 Then
                    If e > LBound(listExtend) Then
                        ReDim listAderouler(1 To 1)
                    End If
                Else
                    ReDim Preserve listAderouler(1 To UBound(listAderouler) + 1)
                End If
                listAderouler(UBound(listAderouler)) = lastLin
                ' on étend
                If e > LBound(listExtend) Then
'If num = 45 Then MsgBox "1" & Join(listAderouler, Chr(10))
'MsgBox "listAderouler:" & Join(listAderouler, Chr(10)) & Chr(10) & Join(newListExend, Chr(10))
'old = newListExend
                    newListExend = eclater(newListExend, listAderouler, 0, nomList, areList, sceList)
'MsgBox "l2 eclater" & Chr(10) & "old=" & Join(old, Chr(10)) & Chr(10) & Chr(10) & "ade=" & Join(listAderouler, Chr(10)) & Chr(10) & Chr(10) & "res=" & Join(newListExend, Chr(10))
'MsgBox "1_____" & Chr(10) & Join(newListExend, Chr(10))
                End If
                ReDim listAderouler(0)
            End If
            lastLin = currlin
'If num = 45 Then MsgBox lastLin & chr(10) a currlin
            lastLev = currLev
        Next
'MsgBox "fin"
        If UBound(listAderouler) = 0 Then
            If e > LBound(listExtend) Then
                ReDim listAderouler(1 To 1)
            End If
        Else
            ReDim Preserve listAderouler(1 To UBound(listAderouler) + 1)
        End If
        listAderouler(UBound(listAderouler)) = lastLin
'If num = 45 Then MsgBox lastLin
'MsgBox "2:" & Join(listAderouler, Chr(10))
'''MsgBox "A1" & Chr(10) & Join(listAderouler, Chr(10)) & Chr(10) & Chr(10) & Join(newListExend, Chr(10))
old = newListExend
        newListExend = eclater(newListExend, listAderouler, 0, nomList, areList, sceList)
'MsgBox UBound(newListExend) & "  2_____" & Chr(10) & Join(newListExend, Chr(10))
'MsgBox "l1 eclater" & Chr(10) & "old=" & Join(old, Chr(10)) & Chr(10) & Chr(10) & "ade=" & Join(listAderouler, Chr(10)) & Chr(10) & Chr(10) & "res=" & Join(newListExend, Chr(10))
        'newListExend = eclater("µ", newListExend, newListExend0, 0, nomList, areList, sceList)
'MsgBox "l2eclater" & Chr(10) & Join(old, Chr(10)) & Chr(10) & Chr(10) & Join(newListExend0, Chr(10)) & Chr(10) & Chr(10) & Join(newListExend, Chr(10))

        lastUbound = UBound(listExtendEnd)
        ReDim Preserve listExtendEnd(LBound(listExtendEnd) To UBound(listExtendEnd) + UBound(newListExend))
        For e = (lastUbound + 1) To UBound(listExtendEnd)
            listExtendEnd(e) = newListExend(e - lastUbound)
        Next
'''MsgBox "B" & Chr(10) & Join(listExtendEnd, Chr(10))
    Next
'MsgBox "___" & Chr(10) & Join(listExtendEnd, Chr(10))
'''If num = 7 Then MsgBox "=Fin=" & Chr(10) & Join(listExtend, Chr(10))
    ' renumérotation
    Dim lastnum As String
    lastnum = ""
    Dim numero As Integer
    numero = 0
    Dim listExtendEndEnd() As String
    ReDim listExtendEndEnd(LBound(listExtendEnd) To UBound(listExtendEnd))
    For e = LBound(listExtendEnd) To UBound(listExtendEnd)
        curnum = Split(listExtendEnd(e) & "+", "+")(0)
        If curnum <> "" Then
            curnum1 = Split(curnum & "_", "_")(1)
            If curnum1 <> "" Then
'MsgBox e & Chr(10) & lastnum & Chr(10) & curnum & Chr(10) & curnum1
                If curnum <> lastnum Then
                    listExtendEndEnd(e) = listExtendEnd(e)
                    numero = 0
                Else
                    If numero = 0 Then
                        numero = CInt(curnum1) + 1
                    Else
                        numero = numero + 1
                    End If
                    listExtendEndEnd(e) = Split(curnum & "_", "_")(0) & "_" & numero & "+" & Split(listExtendEnd(e), "+")(1)
                End If
            Else
                numero = 0
                listExtendEndEnd(e) = listExtendEnd(e)
            End If
        Else
            numero = 0
            listExtendEndEnd(e) = listExtendEnd(e)
        End If
        lastnum = curnum
    Next
'MsgBox "getMatch fin " & Chr(10) & Join(listLin, Chr(10)) & Chr(10) & Chr(10) & Join(listExtendEnd, Chr(10)) & Chr(10) & Chr(10) & Join(listExtendEndEnd, Chr(10))
    ' on vire le | à la fin s'il persiste
    For e = LBound(listExtendEndEnd) To UBound(listExtendEndEnd)
        If Right(listExtendEndEnd(e), 1) = "|" Then
            listExtendEndEnd(e) = Left(listExtendEndEnd(e), Len(listExtendEndEnd(e)) - 1)
        End If
    Next
    getMatch = listExtendEndEnd
    Exit Function
errorHandler: Call onErrDo("Il y a des erreurs", "getMatch"): Exit Function
End Function
Function eclater(newListEx() As String, listAderou() As String, num As Integer, nomList() As String, areList() As String, sceList() As String) As String()
    ' cas d'un item différent du précédent donc on alimente les précédents
    ' attention revoir système de numérotage itération avant le +
    On Error GoTo errorHandler
    Dim ajoutExt() As String
    Dim ok4add As Boolean
    Dim ajout As String
    Dim list() As String
    Dim separateur As String
    separateur = "@"
    ' séparateur niveau 2 @
    splListAder1 = Split(Split(listAderou(1), "+")(1), separateur)
    entete = Split(listAderou(1), "+")(0)
    'If UBound(newListEx) = 0 And Mid(entete, 1, 2) = "2." Then
    'If UBound(newListEx) = 0 Then
'MsgBox entete
        'eclater = newListEx
        'Exit Function
    'End If
    Dim old() As String
    old = newListEx
'MsgBox "avant=" & LBound(newListEx) & ":" & UBound(newListEx) & Chr(10) & Join(newListEx, Chr(10)) & Chr(10) & "adero=" & LBound(listAderou) & ":" & UBound(listAderou) & Chr(10) & Join(listAderou, Chr(10))
    For dup = LBound(splListAder1) To UBound(splListAder1)
        For lad = LBound(listAderou) To UBound(listAderou)
            splListAder = Split(Split(listAderou(lad), "+")(1), separateur)
            If splListAder(dup) <> "" Then
                ok4add = True
                If Left(splListAder(dup), 1) = "[" Then
                    ajout = Split(Mid(splListAder(dup), 2), "]")(0)
                    If Left(ajout, 1) = "s" Then dimension = "SCENARIO"
                    If Left(ajout, 1) = "a" Then dimension = "AREA"
                    If Left(ajout, 1) <> "s" And Left(ajout, 1) <> "a" Then dimension = "NOMENCLATURE"
                    If dimension = "NOMENCLATURE" Then list = nomList
                    If dimension = "AREA" Then list = areList
                    If dimension = "SCENARIO" Then list = sceList
        
                    ajoutExt = getExtended("per", True, ajout, list, num, "QUANTITE", "", "L")
                    If UBound(ajoutExt) = 1 Then
                        ok4add = True
                    Else
                        ok4add = False
                    End If
                End If
                If ok4add Then
                    If UBound(newListEx) = 0 Then
                        ReDim newListEx(1 To 1)
                    Else
                        ReDim Preserve newListEx(1 To UBound(newListEx) + 1)
                    End If
                    splListAder = Split(Split(listAderou(lad), "+")(1), separateur)
                    newListEx(UBound(newListEx)) = Split(listAderou(lad), "+")(0) & "+" & splListAder(dup)
                End If

            End If
        Next
    Next
'MsgBox "apres=" & LBound(newListEx) & ":" & UBound(newListEx) & Chr(10) & Join(newListEx, Chr(10))
    ' separateur niveau 3 µ
    separateur = "µ"
    'splListAder1 = Split(Split(listAderou(1), "+")(1), separateur)
'MsgBox "ICI" & Chr(10) & Join(newListEx, Chr(10)) & Chr(10) & Chr(10) & Join(splListAder1, Chr(10)) & Chr(10) & Chr(10) & Join(listAderou, Chr(10))
    Dim res() As String
    ReDim res(0)
    Dim nbr As Integer
    nbr = 0
    If UBound(newListEx) = 0 Then
        res = newListEx
    Else
    Dim numIter As Integer
    
    For dup = LBound(newListEx) To UBound(newListEx)
        numIter = CInt(Split(Split(newListEx(dup), ".")(1), "-")(0))
        'If Mid(newListEx(dup), 1, 1) = "2" Or Mid(newListEx(dup), 1, 1) = "3" Then
        If numIter >= 2 Then
            listAderou = Split(Split(newListEx(dup), "+")(1), separateur)
            entete = Split(newListEx(dup), "+")(0)
'MsgBox UBound(listAderou) & ">>>>>" & newListEx(dup) & Chr(10) & entete & Chr(10) & Join(listAderou, Chr(10))
            For lad = LBound(listAderou) To UBound(listAderou)
                ok4add = True
                If Left(listAderou(lad), 1) = "[" Then
'MsgBox listAderou(lad)
                    ajout = Split(Mid(listAderou(lad), 2), "]")(0)
                    If Left(ajout, 1) = "s" Then dimension = "SCENARIO"
                    If Left(ajout, 1) = "a" Then dimension = "AREA"
                    If Left(ajout, 1) <> "s" And Left(ajout, 1) <> "a" Then dimension = "NOMENCLATURE"
                    If dimension = "NOMENCLATURE" Then list = nomList
                    If dimension = "AREA" Then list = areList
                    If dimension = "SCENARIO" Then list = sceList
        
                    ajoutExt = getExtended("per", True, ajout, list, num, "QUANTITE", "", "L")
                    If UBound(ajoutExt) = 1 Then
                        ok4add = True
                    Else
                        ok4add = False
                    End If
                End If
                If ok4add Then
                    If UBound(res) = 0 Then
                        ReDim res(1 To 1)
                    Else
                        ReDim Preserve res(1 To UBound(res) + 1)
                    End If
                    nbr = nbr + 1
                    If lad = LBound(listAderou) Then
                        res(nbr) = entete & "+" & listAderou(lad)
                    Else
                            'MsgBox "ici1" & lad & ":" & listAderou(lad)
                                'ReDim Preserve newListEx(LBound(newListEx) To UBound(newListEx) + 1)
                                'MsgBox "ici2" & lad & ":" & UBound(newListEx)
                                'For dup2 = LBound(newListEx) To UBound(newListEx)
                        numlig = CInt(Split(entete, ".")(0))
                        res(nbr) = (numlig + 1) & ".3-1_1" & "+" & listAderou(lad)
                                'Next
                                'MsgBox "ici3" & lad & ":" & UBound(newListEx)
                            'splListAder = Split(Split(listAderou(lad), "+")(1), separateur)
                            'newListEx(UBound(newListEx)) = "3.3" & "+" & splListAder(dup)
                            'MsgBox splListAder(dup)
                    End If
'MsgBox newListEx(dup) & Chr(10) & entete & Chr(10) & Join(listAderou, Chr(10)) & Chr(10) & res(nbr)
                End If
                'End If
            Next
        Else
'MsgBox dup & "d" & UBound(res)
            If UBound(res) = 0 Then
                ReDim res(1 To 1)
            Else
                ReDim Preserve res(1 To UBound(res) + 1)
            End If
            nbr = nbr + 1
            res(nbr) = newListEx(dup)
        End If
    Next
    End If
    eclater = res
'MsgBox Join(eclater, Chr(10))
'MsgBox ">>>>>ava=" & LBound(old) & ":" & UBound(old) & "<>" & Join(old, Chr(10)) & Chr(10) _
'        & ">>>>>ade=" & LBound(listAderou) & ":" & UBound(listAderou) & "<>" & Join(listAderou, Chr(10)) & Chr(10) _
'        & ">>>>>int=" & LBound(newListEx) & ":" & UBound(newListEx) & "<>" & Join(newListEx, Chr(10)) & Chr(10) _
'        & ">>>>>res=" & LBound(eclater) & ":" & UBound(eclater) & "<>" & Join(eclater, Chr(10))
    Exit Function
errorHandler: Call onErrDo("Il y a des erreurs", "eclater"): Exit Function
End Function

Sub SetSheets()
' Lecture des entités de la dimension en feuille
    Dim parDim As String
    parDim = "SCENARIO"
    Dim dimensionCol As Integer
    Dim dimensionLig As Integer
    
    Dim dercol As Integer
    Dim derlig As Integer
    Set FLQUA = Worksheets("QUA")
    Set FLSHEET = Worksheets(parDim)
    Dim sheet() As Variant
    Dim SHEETList() As String
    Dim SHEETListName() As String
    derlignom = FLSHEET.Range("A" & FLSHEET.Rows.Count).End(xlUp).Row
    Valeurs = FLSHEET.Range("A2:B" & derlignom).VALUE
    ReDim SHEETList(1 To UBound(Valeurs, 1))
    ReDim SHEETListName(1 To UBound(Valeurs, 1))
    For a = LBound(Valeurs, 1) To UBound(Valeurs, 1)
        SHEETList(a) = Valeurs(a, 1)
        SHEETListName(a) = Valeurs(a, 2)
    Next
    racineSHEET = SHEETList(1)
    Dim stockSheet() As Variant
    ReDim stockSheet(1 To UBound(SHEETList))
    For I = LBound(stockSheet) To UBound(stockSheet)
        Dim res() As Variant
        ReDim res(1 To 3, 0)
        stockSheet(I) = res
    Next
' Lecture des quantités
    Dim Cla() As Variant
    derlig = FLQUA.Cells.SpecialCells(xlCellTypeLastCell).Row
    dercol = FLQUA.Cells(1, Columns.Count).End(xlToLeft).Column
    With FLQUA.Range("a" & 2 & ":" & DecAlph(dercol) & derlig)
        ReDim Cla(1 To (derlig - 1), 1 To dercol)
        Cla = .VALUE
    End With
    Dim FirstLig() As Variant
    With FLQUA.Range("a" & 1 & ":" & DecAlph(dercol) & 1)
        ReDim FirstLig(1 To 1, 1 To dercol)
        FirstLig = .VALUE
    End With
    ' Détermination des index de colonne pour SCENARIO et AREA et le reste ?
    Dim scenCol As Integer
    Dim areaCol As Integer
    Dim entiCol As Integer
    Dim nameCol As Integer
    For NoCol = 1 To dercol
        If FirstLig(1, NoCol) = "ENTITE" Then entiCol = NoCol
        If FirstLig(1, NoCol) = "AREA" Then areaCol = NoCol
        If FirstLig(1, NoCol) = "SCENARIO" Then scenCol = NoCol
        If FirstLig(1, NoCol) = "NOM" Then nameCol = NoCol
    Next
    If parDim = "SCENARIO" Then
        dimensionCol = scenCol
        dimensionLig = areaCol
    End If
    If parDim = "AREA" Then
        dimensionCol = areaCol
        dimensionLig = scenCol
    End If
' Constitution du tableau de sortie par entité de la dimension feuille
    Dim lastNameQua As String
    lastNameQua = ""
    Dim dimeEnCours As String
    For SheetEnCours = LBound(SHEETList) To UBound(SHEETList)
        Dim ResEnCours() As Variant
        ReDim ResEnCours(1 To 4, 0)
        For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
            dimeEnCours = Trim(Cla(Nolig, dimensionCol))
            If SHEETList(SheetEnCours) = dimeEnCours Then
                If UBound(ResEnCours, 2) > 0 Then ReDim Preserve ResEnCours(1 To UBound(ResEnCours, 1), 1 To UBound(ResEnCours, 2) + 1)
                If UBound(ResEnCours, 2) = 0 Then ReDim ResEnCours(1 To UBound(ResEnCours, 1), 1 To 1)
                If lastNameQua <> Cla(Nolig, nameCol) Then ResEnCours(1, UBound(ResEnCours, 2)) = Cla(Nolig, nameCol)
                ResEnCours(2, UBound(ResEnCours, 2)) = Cla(Nolig, entiCol)
                ResEnCours(3, UBound(ResEnCours, 2)) = Cla(Nolig, dimensionLig)
                ResEnCours(4, UBound(ResEnCours, 2)) = Cla(Nolig, dimensionCol)
                lastNameQua = Cla(Nolig, nameCol)
            End If
        Next
        Set FLFINA = Worksheets("OUT_" & SHEETListName(SheetEnCours))
        FLFINA.Cells.Clear
        If UBound(ResEnCours, 1) > 0 And UBound(ResEnCours, 2) > 0 Then
            Dim ResEnFinal() As Variant
            ReDim ResEnFinal(1 To UBound(ResEnCours, 2), 1 To UBound(ResEnCours, 1))
            If UBound(ResEnCours, 2) > 1 Then
            ResEnFinal = Application.Transpose(ResEnCours)
            FLFINA.Range("A1:" & DecAlph(UBound(ResEnFinal, 2)) & UBound(ResEnFinal, 1)).VALUE = ResEnFinal
            Else
                For c = 1 To UBound(ResEnCours, 1)
                    ResEnFinal(1, c) = ResEnCours(c, 1)
                Next
                FLFINA.Range("A1:" & DecAlph(UBound(ResEnFinal, 2)) & UBound(ResEnFinal, 1)).VALUE = ResEnFinal
            End If
        End If
    Next
End Sub
Function initGlobales() As Boolean
    ACTION = 1
    ID = 2
    BEFORE1 = 3
    THIS1 = 4
    ATTR1 = 5
    BEFORE2 = 6
    THIS2 = 7
    ATTR2 = 8
    BEFORE3 = 9
    THIS3 = 10
    ATTR3 = 11
    BEFORE4 = 12
    THIS4 = 13
    ATTR4 = 14
    BEFORE5 = 15
    THIS5 = 16
    ATTR5 = 17
    BEFORE6 = 18
    THIS6 = 19
    ATTR6 = 20
    BEFORE7 = 21
    THIS7 = 22
    ATTR7 = 23
    BEFORE8 = 24
    THIS8 = 25
    ATTR8 = 26
    BEFORE9 = 27
    THIS9 = 28
    ATTR9 = 29
    EXTENDED = 30
    NAME = 31
    AREA = 32
    SYNONYMOUS = 32
    SCENARIO = 33
    quantity = 34
    temps = 37
    equation = 38
    ''SCALE   UNIT    TIME    FORMULA SUBSTITUTE  DEFAUT  VALUE   BOUCLE
    SCALE0 = 35
    UNIT = 36
    SUBSTITUTE = 39
    DEFAUT = 40
    VALUE = 41
    BOUCLE = 42
    initGlobales = True
End Function
Function coloriage(ligDeb As Integer, FL0 As Worksheet, FLC As Worksheet) As Integer
    ' Coloriage des alertes et erreurs
    derligsheet = getDerLig(FL0)
    FL0.Range("a" & ligDeb & ":a" & derligsheet).Interior.Color = xlNone
    FL0.Rows(ligDeb & ":" & derligsheet).Borders.LineStyle = xlContinuous
    FLC.Rows("1:" & derligsheet).HorizontalAlignment = xlLeft
    Dim nberror As Integer
    nberror = 0
    derligsheet = FLC.Cells.SpecialCells(xlCellTypeLastCell).Row
    For Nolig = 1 To derligsheet
        col = "a"
        elt = Trim(FLC.Cells(Nolig, 6))
        If Left(Trim(FL0.Cells(Nolig, 1)), 1) = "#" Then
            '''FL0.Range("a" & Nolig & ":ae" & Nolig).Interior.Color = RGB(230, 230, 230)
            FL0.Range("a" & Nolig & ":a" & Nolig).Interior.Color = RGB(230, 230, 230)
        End If
        splElt = Split(elt, ",")
        For I = LBound(splElt) To UBound(splElt)
            If splElt(I) = "ID" Then col = col & "," & "B"
            If splElt(I) = "BEFORE1" Then col = col & "," & "C"
            If splElt(I) = "THIS1" Then col = col & "," & "D"
            If splElt(I) = "ATTR1" Then col = col & "," & "E"
            If splElt(I) = "BEFORE2" Then col = col & "," & "F"
            If splElt(I) = "THIS2" Then col = col & "," & "G"
            If splElt(I) = "ATTR2" Then col = col & "," & "H"
            If splElt(I) = "BEFORE3" Then col = col & "," & "I"
            If splElt(I) = "THIS3" Then col = col & "," & "J"
            If splElt(I) = "ATTR3" Then col = col & "," & "K"
            If splElt(I) = "BEFORE4" Then col = col & "," & "L"
            If splElt(I) = "THIS4" Then col = col & "," & "M"
            If splElt(I) = "ATTR4" Then col = col & "," & "N"
            If splElt(I) = "BEFORE5" Then col = col & "," & "O"
            If splElt(I) = "THIS5" Then col = col & "," & "P"
            If splElt(I) = "ATTR5" Then col = col & "," & "Q"
            If splElt(I) = "BEFORE6" Then col = col & "," & "R"
            If splElt(I) = "THIS6" Then col = col & "," & "S"
            If splElt(I) = "ATTR6" Then col = col & "," & "T"
            If splElt(I) = "BEFORE7" Then col = col & "," & "U"
            If splElt(I) = "THIS7" Then col = col & "," & "V"
            If splElt(I) = "ATTR7" Then col = col & "," & "W"
            If splElt(I) = "BEFORE8" Then col = col & "," & "X"
            If splElt(I) = "THIS8" Then col = col & "," & "Y"
            If splElt(I) = "ATTR8" Then col = col & "," & "z"
            If splElt(I) = "BEFORE9" Then col = col & "," & "aa"
            If splElt(I) = "THIS9" Then col = col & "," & "ab"
            If splElt(I) = "ATTR9" Then col = col & "," & "ac"
            If splElt(I) = "AREA" Then col = col & "," & "af"
            If splElt(I) = "SCENARIO" Then col = col & "," & "ag"
            If splElt(I) = "TEMPS" Then col = col & "," & "ak"
            If splElt(I) = "EQUATION" Then col = col & "," & "al"
        Next
        splCol = Split(col, ",")
        If FLC.Cells(Nolig, 2) = "DEBUT" Then
            FLC.Range("a" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 255, 153)
        End If
        If FLC.Cells(Nolig, 2) = "FIN" Then
            FLC.Range("a" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 255, 153)
        End If
        If FLC.Cells(Nolig, 2) = "INFO" Then
            FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 204, 255)
        End If
        If FLC.Cells(Nolig, 2) Like "ERREUR*" Then
            nberror = nberror + 1
            If FLC.Cells(Nolig, 2) = "ERREUR" Then
                FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(255, 160, 160)
            Else
                FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(255, 165, 0)
            End If
            '''If f = t Then
                For I = LBound(splCol) To UBound(splCol)
                    If splCol(I) = "a" Then
                        If FLC.Cells(Nolig, 2) = "ERREUR" Then
                            If (InStr(FLC.Cells(Nolig, 8).VALUE, ";") > 0) Or (FLC.Cells(Nolig, 8).VALUE = "") Then
                            
                            Else
                                FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Interior.ColorIndex = 3
                            End If
                        Else
                            FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Interior.ColorIndex = 46
                        End If
                    Else
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.LineStyle = xlContinuous
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.Weight = xlThick
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.ColorIndex = 3
                    End If
                Next
            '''End If
        End If
        If FLC.Cells(Nolig, 2) = "ALERTE" Then
            FLC.Range("b" & Nolig).Interior.ColorIndex = 44
            'FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(255, 230, 100)
            If Trim(FLC.Cells(Nolig, 8).VALUE) <> "" Then
                For I = LBound(splCol) To UBound(splCol)
                    '''FL0.Range(splCol(i) & FLC.Cells(Nolig, 8).Value).Interior.Color = RGB(255, 230, 100)
                    If splCol(I) = "a" Then
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Interior.ColorIndex = 45
                    Else
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.LineStyle = xlContinuous
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.Weight = xlThick
                        FL0.Range(splCol(I) & FLC.Cells(Nolig, 8).VALUE).Cells.Borders.ColorIndex = 45
                    End If
                Next
            End If
        End If
    Next
    coloriage = nberror
End Function
Function coloriageLog(FLC As Worksheet) As Integer
    ' Coloriage des infos, alertes et erreurs dans le log FLC
    Dim derligsheet As Integer
    derligsheet = getDerLig(FLC)
    FLC.Rows("1:" & derligsheet).HorizontalAlignment = xlLeft
    FLC.Range("A1:K" & UBound(logNom, 1)).Cells.Borders.LineStyle = xlContinuous
    Dim nberror As Integer
    nberror = 0
    derligsheet = FLC.Cells.SpecialCells(xlCellTypeLastCell).Row
    For Nolig = 1 To derligsheet
        If FLC.Cells(Nolig, 2) = "DEBUT" Then
            FLC.Range("a" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 255, 153)
        End If
        If FLC.Cells(Nolig, 2) = "FIN" Then
            FLC.Range("a" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 255, 153)
        End If
        If FLC.Cells(Nolig, 2) = "INFO" Then
            FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(153, 204, 255)
        End If
        If FLC.Cells(Nolig, 2) Like "ERREUR*" Then
            nberror = nberror + 1
            If FLC.Cells(Nolig, 2) = "ERREUR" Then
                FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(255, 160, 160)
            Else
                FLC.Range("b" & Nolig & ":i" & Nolig).Interior.Color = RGB(255, 165, 0)
            End If
        End If
        If FLC.Cells(Nolig, 2) = "ALERTE" Then
            FLC.Range("b" & Nolig).Interior.ColorIndex = 44
        End If
    Next
    coloriageLog = nberror
End Function

Sub SetNomenclatureNew()
    Dim sngChrono As Single
    sngChrono = Timer
    Dim FL0NO As Worksheet
    Dim FLCNO As Worksheet
    Set FL0NO = Worksheets(ActiveSheet.NAME)
    Set FLCNO = Worksheets("CRNOM")
    Set FLNOM = Worksheets("NOM")
    FLNOM.Cells.Clear
    resInitGlobales = initGlobales()
    Dim LigDeb0No As Integer
    'Call SetFl              ' Initialisation des feuilles
    'Call DelNomenclature    ' Initialisation de la feuille NOMENCLATURE
    'FL0NO.Columns(EXTENDED).EntireColumn.Delete
    Dim ATTRIBUTS() As String
    Dim simples() As String
    Dim simplesNoAt() As String
    Dim sets() As String
    Dim all() As String
    Dim numlig() As Integer
    Dim extendNb() As Integer
    
    derlig = Split(FLCNO.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLCNO.Cells(Nolig, 1) = "CONTEXTE" Then
            LigDebCNo = Nolig + 1
            Exit For
        End If
    Next
    derligsheet = FLCNO.Range("A" & FLCNO.Rows.Count).End(xlUp).Row
    DerColSheet = Cells(LigDebCNo - 1, FLCNO.Columns.Count).End(xlToLeft).Column
    FLCNO.Range("A" & LigDebCNo & ":" & "K" & WorksheetFunction.Max(LigDebCNo, derligsheet)).Clear
    With FLCNO.Range("a" & (LigDebCNo - 1) & ":" & "K" & (LigDebCNo - 1))
        ReDim logNom(1 To DerColSheet, 1 To (LigDebCNo - 1))
        logNom = Application.Transpose(.VALUE)
    End With
    Dim newLogLine() As Variant
    newLogLine = Array("NOMENCLATURE", "", "0NOM", "GENERATION", "DEBUT", time())
    logNom = alimLog(logNom, newLogLine)

    'DerLig = Split(FL0NO.UsedRange.Address, "$")(4)
    'Dim Ligne As Long
    derlig = FL0NO.Cells.SpecialCells(xlCellTypeLastCell).Row
    For Nolig = 1 To derlig
        If FL0NO.Cells(Nolig, 1) = "ACTION" Then
            LigDeb0No = Nolig + 1
            Exit For
        End If
    Next
    'FL0NO.Cells(LigDeb0No - 1, EXTENDED).Value = "EXTENDED"

    'Alimentation des tableaux des classes et autres
    Dim Cla() As Variant

    'DerLigSheet = FL0NO.Range("A" & Rows.Count).End(xlUp).Row
    derligsheet = FL0NO.Cells.SpecialCells(xlCellTypeLastCell).Row
    DerColSheet = Cells(LigDeb0No - 1, Columns.Count).End(xlToLeft).Column

    With FL0NO.Range("a" & LigDeb0No & ":" & "AC" & derligsheet)
        ReDim Cla(1 To (derligsheet - LigDeb0No), 1 To DerColSheet)
        Cla = .VALUE
    End With

    ' premier passage pour déterminer les dimensions
    Dim DimAttributs As Integer, DimSimples As Integer, DimSets As Integer, DimAll As Integer
    DimAttributs = 0
    DimSimples = 0
    DimSets = 0
    DimAll = 0
    Dim vecBool() As Boolean
    ReDim vecBool(LBound(Cla, 1) To UBound(Cla, 1))
    Dim vecteur As String
    Dim numeroDeLigne As Integer
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
        vecBool(Nolig) = True
        vecteur = Cla(Nolig, ACTION) & "@" & Cla(Nolig, ID) & "@" & Cla(Nolig, BEFORE1) & "@" & Cla(Nolig, THIS1) & "@" & Cla(Nolig, ATTR1) _
        & "@" & Cla(Nolig, BEFORE2) & "@" & Cla(Nolig, THIS2) & "@" & Cla(Nolig, ATTR2) _
        & "@" & Cla(Nolig, BEFORE3) & "@" & Cla(Nolig, THIS3) & "@" & Cla(Nolig, ATTR3) _
        & "@" & Cla(Nolig, BEFORE4) & "@" & Cla(Nolig, THIS4) & "@" & Cla(Nolig, ATTR4) _
        & "@" & Cla(Nolig, BEFORE5) & "@" & Cla(Nolig, THIS5) & "@" & Cla(Nolig, ATTR5) _
        & "@" & Cla(Nolig, BEFORE6) & "@" & Cla(Nolig, THIS6) & "@" & Cla(Nolig, ATTR6) _
        & "@" & Cla(Nolig, BEFORE7) & "@" & Cla(Nolig, THIS7) & "@" & Cla(Nolig, ATTR7) _
        & "@" & Cla(Nolig, BEFORE8) & "@" & Cla(Nolig, THIS8) & "@" & Cla(Nolig, ATTR8) _
        & "@" & Cla(Nolig, BEFORE9) & "@" & Cla(Nolig, THIS9) & "@" & Cla(Nolig, ATTR9)
        If Left(Cla(Nolig, 1), 1) <> "#" Then
            numeroDeLigne = LigDeb0No + Nolig - 1
            resTestForme = testForme(vecteur, numeroDeLigne)
            If Not resTestForme Then
                vecBool(Nolig) = False
                GoTo ContinueLoop1
            End If
            If Cla(Nolig, BEFORE1) = "" And Cla(Nolig, THIS1) = "" And Cla(Nolig, ATTR1) <> "" Then
                DimAttributs = DimAttributs + 1
            End If
            If (Cla(Nolig, BEFORE1) <> "" Or Cla(Nolig, THIS1) <> "") And Cla(Nolig, BEFORE2) = "" And Cla(Nolig, THIS2) = "" Then
                DimSimples = DimSimples + 1
            End If
            If Cla(Nolig, BEFORE1) <> "" And Cla(Nolig, THIS1) = "" And Cla(Nolig, BEFORE2) = "" And Cla(Nolig, THIS2) = "" Then
                DimSets = DimSets + 1
            End If
            If Cla(Nolig, BEFORE1) <> "" Or Cla(Nolig, THIS1) <> "" Then
                DimAll = DimAll + 1
            End If
        End If
ContinueLoop1:
    Next
    newLogLine = Array("", "INFO", "0NOM", "", "", time(), "", "", DimAttributs, "", "Nombre d'attributs")
    logNom = alimLog(logNom, newLogLine)
    'newLogLine = Array("", "INFO", "0NOM", "", "", Time(), "", "", DimSets, "", "Nombre d'ensembles")
    'logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("", "INFO", "0NOM", "", "", time(), "", "", DimSimples, "", "Nombre d'entités simples")
    logNom = alimLog(logNom, newLogLine)
    If DimAttributs > 0 Then
        ReDim ATTRIBUTS(1 To DimAttributs)
    Else
        ReDim ATTRIBUTS(0 To DimAttributs)
    End If
    ReDim simples(1 To DimSimples) ', 1 To 1)
    ReDim simplesNoAt(1 To DimSimples) ', 1 To 1)
    ReDim sets(1 To DimSets)
    ReDim all(1 To DimAll, 1 To 3)
    ReDim numlig(1 To DimAll)
    'ReDim extendNb(1 To DimAll)
    Dim posAttr As Integer, posSimp As Integer, posSets As Integer, posAll As Integer
    posAttr = 0
    posSimp = 0
    posSets = 0
    posAll = 0
    Dim jonction As String, jonction1 As String, jonction2 As String, jonction3 As String, jonction4 As String, jonction5 As String, jonction6 As String, jonction7 As String, jonction8 As String, jonction9 As String
    Dim attj As String, attj1 As String, attj2 As String, attj3 As String, attj4 As String, attj5 As String, attj6 As String, attj7 As String, attj8 As String, attj9 As String
    Dim point12 As String, point23 As String, point34 As String, point45 As String, point56 As String, point67 As String, point78 As String, point89 As String
    
    ' passage 2 pour alimenter les tableaux
    For Nolig = LBound(Cla, 1) To UBound(Cla, 1)
    If vecBool(Nolig) Then
        If Left(Cla(Nolig, 1), 1) <> "#" Then
            ' traitement des attributs
            If Cla(Nolig, BEFORE1) = "" And Cla(Nolig, THIS1) = "" And Cla(Nolig, ATTR1) <> "" Then
                posAttr = posAttr + 1
                ATTRIBUTS(posAttr) = Cla(Nolig, ATTR1)
            End If
            If (Cla(Nolig, BEFORE1) <> "" Or Cla(Nolig, THIS1) <> "") And Cla(Nolig, BEFORE2) = "" And Cla(Nolig, THIS2) = "" Then
                posSimp = posSimp + 1
                If Cla(Nolig, THIS1) = "" Then jonction = ""
                If Cla(Nolig, THIS1) <> "" Then jonction = ">"
                If Cla(Nolig, ATTR1) = "" Then attj = ""
                'If Cla(Nolig, ATTR1) <> "" Then attj = ":"

                If Trim(Cla(Nolig, ATTR1)) <> "" Then
lat = isAnAttr(delSpace(Trim(Cla(Nolig, ATTR1))), ATTRIBUTS)
If lat <> "" Then
    newLogLine = Array("", "ALERTE", "0NOM", "", time(), "", "ATTR1", lat, LigDeb0No + Nolig - 1, "attribut(s) non défini(s)")
    logNom = alimLog(logNom, newLogLine)
    vecBool(Nolig) = False
End If
                    attj = ":"
                End If
                simples(posSimp) = nettoyage(Cla(Nolig, BEFORE1) & jonction & Cla(Nolig, THIS1) & attj & Cla(Nolig, ATTR1))
                '''simplesNoAt(posSimp) = nettoyage(Cla(Nolig, BEFORE1) & jonction & Cla(Nolig, THIS1))
            End If
            If Cla(Nolig, BEFORE1) <> "" And Cla(Nolig, THIS1) = "" And Cla(Nolig, BEFORE2) = "" And Cla(Nolig, THIS2) = "" Then
                posSets = posSets + 1
                If Cla(Nolig, ATTR1) = "" Then attj = ""
                If Trim(Cla(Nolig, ATTR1)) <> "" Then attj = ":"
                sets(posSets) = nettoyage(Cla(Nolig, BEFORE1) & attj & Cla(Nolig, ATTR1))
            End If
            If Cla(Nolig, BEFORE1) <> "" Or Cla(Nolig, THIS1) <> "" Then
                posAll = posAll + 1
                If Cla(Nolig, THIS1) = "" Then jonction1 = ""
                If Cla(Nolig, THIS1) <> "" Then jonction1 = ">"
                If Cla(Nolig, THIS2) = "" Then jonction2 = ""
                If Cla(Nolig, THIS2) <> "" Then jonction2 = ">"
                If Cla(Nolig, THIS3) = "" Then jonction3 = ""
                If Cla(Nolig, THIS3) <> "" Then jonction3 = ">"
                If Cla(Nolig, THIS4) = "" Then jonction4 = ""
                If Cla(Nolig, THIS4) <> "" Then jonction4 = ">"
                If Cla(Nolig, THIS5) = "" Then jonction5 = ""
                If Cla(Nolig, THIS5) <> "" Then jonction5 = ">"
                If Cla(Nolig, THIS6) = "" Then jonction6 = ""
                If Cla(Nolig, THIS6) <> "" Then jonction6 = ">"
                If Cla(Nolig, THIS7) = "" Then jonction7 = ""
                If Cla(Nolig, THIS7) <> "" Then jonction7 = ">"
                If Cla(Nolig, THIS8) = "" Then jonction8 = ""
                If Cla(Nolig, THIS8) <> "" Then jonction8 = ">"
                If Cla(Nolig, THIS9) = "" Then jonction9 = ""
                If Cla(Nolig, THIS9) <> "" Then jonction9 = ">"
                If Cla(Nolig, BEFORE2) <> "" Or Cla(Nolig, THIS2) <> "" Then
                    point12 = "."
                Else
                    point12 = ""
                End If
                If Cla(Nolig, BEFORE3) <> "" Or Cla(Nolig, THIS3) <> "" Then
                    point23 = "."
                Else
                    point23 = ""
                End If
                If Cla(Nolig, BEFORE4) <> "" Or Cla(Nolig, THIS4) <> "" Then
                    point34 = "."
                Else
                    point34 = ""
                End If
                If Cla(Nolig, BEFORE5) <> "" Or Cla(Nolig, THIS5) <> "" Then
                    point45 = "."
                Else
                    point45 = ""
                End If
                If Cla(Nolig, BEFORE6) <> "" Or Cla(Nolig, THIS6) <> "" Then
                    point56 = "."
                Else
                    point56 = ""
                End If
                If Cla(Nolig, BEFORE7) <> "" Or Cla(Nolig, THIS7) <> "" Then
                    point67 = "."
                Else
                    point67 = ""
                End If
                If Cla(Nolig, BEFORE8) <> "" Or Cla(Nolig, THIS8) <> "" Then
                    point78 = "."
                Else
                    point78 = ""
                End If
                If Cla(Nolig, BEFORE9) <> "" Or Cla(Nolig, THIS9) <> "" Then
                    point89 = "."
                Else
                    point89 = ""
                End If
                If Trim(Cla(Nolig, ATTR1)) = "" Then attj1 = ""
                If Trim(Cla(Nolig, ATTR1)) <> "" Then attj1 = ":"
                If Trim(Cla(Nolig, ATTR2)) = "" Then attj2 = ""
                If Trim(Cla(Nolig, ATTR2)) <> "" Then attj2 = ":"
                If Trim(Cla(Nolig, ATTR3)) = "" Then attj3 = ""
                If Trim(Cla(Nolig, ATTR3)) <> "" Then attj3 = ":"
                If Trim(Cla(Nolig, ATTR4)) = "" Then attj4 = ""
                If Trim(Cla(Nolig, ATTR4)) <> "" Then attj4 = ":"
                If Trim(Cla(Nolig, ATTR5)) = "" Then attj5 = ""
                If Trim(Cla(Nolig, ATTR5)) <> "" Then attj5 = ":"
                If Trim(Cla(Nolig, ATTR6)) = "" Then attj6 = ""
                If Trim(Cla(Nolig, ATTR6)) <> "" Then attj6 = ":"
                If Trim(Cla(Nolig, ATTR7)) = "" Then attj7 = ""
                If Trim(Cla(Nolig, ATTR7)) <> "" Then attj7 = ":"
                If Trim(Cla(Nolig, ATTR8)) = "" Then attj8 = ""
                If Trim(Cla(Nolig, ATTR8)) <> "" Then attj8 = ":"
                If Trim(Cla(Nolig, ATTR9)) = "" Then attj9 = ""
                If Trim(Cla(Nolig, ATTR9)) <> "" Then attj9 = ":"
                all(posAll, 1) = nettoyage(Cla(Nolig, BEFORE1) & jonction1 & Cla(Nolig, THIS1) & attj1 & Cla(Nolig, ATTR1) _
                & point12 & Cla(Nolig, BEFORE2) & jonction2 & Cla(Nolig, THIS2) & attj2 & Cla(Nolig, ATTR2) _
                & point23 & Cla(Nolig, BEFORE3) & jonction3 & Cla(Nolig, THIS3) & attj3 & Cla(Nolig, ATTR3) _
                & point34 & Cla(Nolig, BEFORE4) & jonction4 & Cla(Nolig, THIS4) & attj4 & Cla(Nolig, ATTR4) _
                & point45 & Cla(Nolig, BEFORE5) & jonction5 & Cla(Nolig, THIS5) & attj5 & Cla(Nolig, ATTR5) _
                & point56 & Cla(Nolig, BEFORE6) & jonction6 & Cla(Nolig, THIS6) & attj6 & Cla(Nolig, ATTR6) _
                & point67 & Cla(Nolig, BEFORE7) & jonction7 & Cla(Nolig, THIS7) & attj7 & Cla(Nolig, ATTR7) _
                & point78 & Cla(Nolig, BEFORE8) & jonction8 & Cla(Nolig, THIS8) & attj8 & Cla(Nolig, ATTR8) _
                & point89 & Cla(Nolig, BEFORE9) & jonction9 & Cla(Nolig, THIS9) & attj9 & Cla(Nolig, ATTR9))
                all(posAll, 2) = nettoyage("" & Cla(Nolig, ATTR1))
                numlig(posAll) = Nolig
                all(posAll, 3) = Trim(Cla(Nolig, ACTION))
            End If
        End If
    End If
    Next
    ' résolution des raccourcis
    Dim entree() As String
    Dim listAction() As String
    Dim sortie As String
    Dim limite As Integer
    Dim newNumLig() As Integer
    limite = 0
    Dim pos As String
    pos = "BEFORE"
    For NoOne = LBound(all, 1) To UBound(all, 1)
        sortie = remplaceDoubleSup(all(NoOne, 1), entree)
        SplSortie = Split(sortie, ".")
        If sortie Like "*>>*" Then
            For I = LBound(SplSortie) To UBound(SplSortie)
                If SplSortie(I) Like "*>>*" Then
                pos = "BEFORE" & (1 + I)
newLogLine = Array("", "ERREUR", "0NOM", "", time(), "", pos, SplSortie(I), LigDeb0No + numlig(NoOne) - 1, "raccourci non résolu")
logNom = alimLog(logNom, newLogLine)
                End If
            Next
        Else
            limite = limite + 1
            ReDim Preserve entree(1 To limite)
            ReDim Preserve newNumLig(1 To limite)
            ReDim Preserve listAction(1 To limite)
            listAction(limite) = all(NoOne, 3)
            entree(limite) = sortie
            newNumLig(limite) = numlig(NoOne)
        End If
    Next
    ' propagation des attributs
    Dim joinSpl() As String
    Dim ind As Integer
    Dim reg As VBScript_RegExp_55.RegExp
    Set reg = New VBScript_RegExp_55.RegExp
    reg.Global = False
    reg.Pattern = ">"
    Dim jcta As String
    Dim aj As String
    For Nolig = LBound(entree) To UBound(entree)
        If entree(Nolig) Like "*.*" Then
            spl = Split(entree(Nolig), ".")
            ReDim joinSpl(UBound(spl))
            ind = 0
            For NoF = LBound(spl) To UBound(spl)
                joinSpl(NoF) = spl(NoF)
            Next
            For Each facteur In spl
                ' on va chercher les attributs
                achercher = Split(facteur, ":")(0)
                Dim attr As String
                attr = ""
                For nol = LBound(entree) To Nolig - 1
                    If Split(entree(nol), ":")(0) = achercher Then
                        If UBound(Split(entree(nol), ":")) > 0 Then attr = Split(entree(nol), ":")(1)
                        Exit For
                    End If
                Next
                If UBound(Split(facteur, ":")) > 0 Then
                    ' cas ou il y a filtrage
                    ' RIEN A FAIRE ?
                Else
                    ' on pose les attributs
                    If attr <> "" Then joinSpl(ind) = facteur & ":" & attr '''all(Nol, 2)
                End If
                ind = ind + 1
            Next
            entree(Nolig) = Join(joinSpl, ".")
        Else
            ent = Split(entree(Nolig), ":")(0)
            If UBound(Split(entree(Nolig), ":")) > 0 Then
                ' Cas de surcharge des attributs
            Else
                entR = StrReverse(Trim(ent))
                If InStr(entR, ">") > 0 Then
                    achercher = Left(ent, Len(ent) - InStr(entR, ">"))
                    For nol = LBound(entree) To Nolig - 1
                        If Split(entree(nol), ":")(0) = achercher Then
                            If UBound(Split(entree(nol), ":")) < 1 Then
                                jcta = ""
                                aj = ""
                            Else
                                jcta = ":"
                                aj = Trim(Split(entree(nol), ":")(1))
                            End If
                            entree(Nolig) = entree(Nolig) & jcta & aj
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' alimentation des simples
    '''For Nolig = LBound(entree) To UBound(entree)
        '''If Not entree(Nolig) Like "*.*" Then
            '''simplesNoAt(Nolig) = Split(entree(Nolig), ":")(0)
        '''End If
    '''Next
    ' test de pédéfinition de l'antécédent
    For e = LBound(entree) To UBound(entree)
        If Not entree(e) Like "*.*" Then
            If Not getPre(entree(e), entree) Then
newLogLine = Array("", "ALERTE", "0NOM", "", time(), "", "BEFORE1", entree(e), LigDeb0No + newNumLig(e) - 1, "antécédent non défini")
logNom = alimLog(logNom, newLogLine)
            End If
        End If
    Next
    ' Extension des produits cartésien
    Dim allVec() As String
    Dim allLig() As String
    Dim allAct() As String
    ReDim allVec(1 To 1)
    ReDim allLig(1 To 1)
    ReDim allAct(1 To 1)
    Dim regEACH As VBScript_RegExp_55.RegExp
    Set regEACH = New VBScript_RegExp_55.RegExp
    Dim newProduit As String
    Dim jct As String
    Dim eten() As String
    ReDim extendNb(1 To UBound(newNumLig))
    For Nolig = LBound(entree) To UBound(entree)
        If entree(Nolig) Like "*.*" Then
            eten = getExtended(False, entree(Nolig), entree, LigDeb0No - 1 + newNumLig(Nolig), "L")
            If UBound(eten) > 0 Then
                If testExtend(entree(Nolig), eten, LigDeb0No - 1 + newNumLig(Nolig)) Then GoTo continueLoop
                Last = UBound(allVec)
                If Nolig > 1 Then
                    ReDim Preserve allVec(LBound(allVec) To (UBound(allVec) + UBound(eten)))
                    ReDim Preserve allLig(LBound(allLig) To (UBound(allLig) + UBound(eten)))
                    ReDim Preserve allAct(LBound(allAct) To (UBound(allAct) + UBound(eten)))
                End If
                For n = LBound(eten) To UBound(eten)
                    allVec(Last + n) = eten(n)
                    allLig(Last + n) = LigDeb0No - 1 + newNumLig(Nolig)
                    allAct(Last + n) = listAction(Nolig)
                Next
                extendNb(Nolig) = UBound(eten)
            Else
            
            End If
        Else
            If Nolig > 1 Then
                ReDim Preserve allVec(LBound(allVec) To (UBound(allVec) + 1))
                ReDim Preserve allLig(LBound(allLig) To (UBound(allLig) + 1))
                ReDim Preserve allAct(LBound(allAct) To (UBound(allAct) + 1))
            End If
            allVec(UBound(allVec)) = entree(Nolig)
            allLig(UBound(allLig)) = LigDeb0No - 1 + newNumLig(Nolig)
            allAct(UBound(allAct)) = listAction(Nolig)
            extendNb(Nolig) = 1
        End If
continueLoop:
    Next

    ' Surcharges
    Dim NbSurcharges As Integer
    NbSurcharges = 0
    Dim LigSurcharges() As String
    Dim surcharge As Boolean
    Dim allVecSur() As String
    ReDim allVecSur(1 To 1)
    allVecSur(1) = allVec(1)
    Dim allLigSur() As String
    ReDim allLigSur(1 To 1)
    allLigSur(1) = allLig(1)
    Dim allActSur() As String
    ReDim allActSur(1 To 1)
    allActSur(1) = allAct(1)
    For Nolig = (LBound(allVec) + 1) To UBound(allVec)
        If allAct(Nolig) <> "minus" Then
        If Not allVec(Nolig) Like "*.*" Then
            surcharge = False
            For nol = LBound(allVecSur) To UBound(allVecSur)
                If allAct(Nolig) <> "minus" Then
                s1 = Split(allVec(Nolig), ":")
                s2 = Split(allVecSur(nol), ":")
                If s1(0) = s2(0) Then
                    allVecSur(nol) = allVec(Nolig)
                    allLigSur(nol) = allLigSur(nol) & "," & allLig(Nolig)
                    surcharge = True
                    NbSurcharges = NbSurcharges + 1
                    ReDim Preserve LigSurcharges(1 To NbSurcharges)
                    LigSurcharges(UBound(LigSurcharges)) = allLig(Nolig)
                    Exit For
                End If
                End If
            Next
            If Not surcharge Then
                ReDim Preserve allVecSur(LBound(allVecSur) To (UBound(allVecSur) + 1))
                allVecSur(UBound(allVecSur)) = allVec(Nolig)
                ReDim Preserve allLigSur(LBound(allLigSur) To (UBound(allLigSur) + 1))
                allLigSur(UBound(allLigSur)) = allLig(Nolig)
                ReDim Preserve allActSur(LBound(allActSur) To (UBound(allActSur) + 1))
                allActSur(UBound(allActSur)) = allAct(Nolig)
            End If
        Else
' A FAIRE surcharges des complexes
            ReDim Preserve allVecSur(LBound(allVecSur) To (UBound(allVecSur) + 1))
            allVecSur(UBound(allVecSur)) = allVec(Nolig)
            ReDim Preserve allLigSur(LBound(allLigSur) To (UBound(allLigSur) + 1))
            allLigSur(UBound(allLigSur)) = allLig(Nolig)
            ReDim Preserve allActSur(LBound(allActSur) To (UBound(allActSur) + 1))
            allActSur(UBound(allActSur)) = allAct(Nolig)
        End If
        Else
        End If
    Next

    ' Minus
    Dim listMinus() As String
    ReDim listMinus(0 To 0)
    Dim listBoolMinus() As Boolean
    ReDim listBoolMinus(0 To 0)
    Dim ligMinus() As String
    ReDim ligMinus(1 To 1)
    Dim cibMinus() As String
    ReDim cibMinus(1 To 1)
    Dim first As Boolean
    first = False
    Dim initMinus As Integer
    For Nolig = LBound(allVec) To UBound(allVec)
        If allAct(Nolig) = "minus" Then
            If first Then
                ReDim Preserve listMinus(1 To (UBound(listMinus) + 1))
                ReDim Preserve listBoolMinus(1 To (UBound(listBoolMinus) + 1))
            Else
                ReDim listMinus(1 To 1)
                ReDim listBoolMinus(1 To 1)
            End If
            If first Then ReDim Preserve ligMinus(1 To (UBound(ligMinus) + 1))
            listMinus(UBound(listMinus)) = allVec(Nolig)
            ligMinus(UBound(listMinus)) = allLig(Nolig)
            listBoolMinus(UBound(listBoolMinus)) = False
            first = True
        End If
    Next

    Dim allVecRet() As String
    ReDim allVecRet(1 To 1)
    'allVecRet(1) = allVec(1)
    Dim allLigRet() As String
    ReDim allLigRet(1 To 1)
    'allLigRet(1) = allLig(1)
    Dim retirer As Boolean
    first = False
    Dim numMinus As Integer

    For Nolig = LBound(allVecSur) To UBound(allVecSur)
        retirer = False
        numMinus = -1
        For aret = LBound(listMinus) To UBound(listMinus)
            If allVecSur(Nolig) = listMinus(aret) Then
                retirer = True
                numMinus = aret
                listBoolMinus(numMinus) = True
                Exit For
            End If
        Next
        If Not retirer Then
            If first Then ReDim Preserve allVecRet(LBound(allVecRet) To (UBound(allVecRet) + 1))
            allVecRet(UBound(allVecRet)) = allVecSur(Nolig)
            If first Then ReDim Preserve allLigRet(LBound(allLigRet) To (UBound(allLigRet) + 1))
            allLigRet(UBound(allLigRet)) = allLigSur(Nolig)
            first = True
        Else
            'listBoolMinus(numMinus) = True
        End If
    Next
    If UBound(listMinus) > 0 Then
        For m = LBound(listBoolMinus) To UBound(listBoolMinus)
            If Not (listBoolMinus(m)) Then
newLogLine = Array("", "ALERTE", "0NOM", "", time(), "", "BEFORE1", listMinus(m), ligMinus(m), "Retrait non résolu")
logNom = alimLog(logNom, newLogLine)
            End If
        Next
    End If
    newLogLine = Array("", "INFO", "0NOM", "", time(), "", "", UBound(listMinus), Join(ligMinus, ","), "Retrait(s)")
    logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("", "INFO", "0NOM", "", time(), "", "", NbSurcharges, Join(LigSurcharges, ","), "Surcharge(s)")
    logNom = alimLog(logNom, newLogLine)
    newLogLine = Array("", "INFO", "0NOM", "", time(), "", "", UBound(entree), UBound(allVecSur), "Nombre d'entités générées")
    logNom = alimLog(logNom, newLogLine)
    
    Dim aecrire() As String
    ReDim aecrire(1 To UBound(allVecRet), 1 To 3)
    For Nolig = 1 To UBound(allVecRet)
        aecrire(Nolig, 1) = FL0NO.NAME
        aecrire(Nolig, 2) = allLigRet(Nolig)
        aecrire(Nolig, 3) = allVecRet(Nolig)
        'aEcrire(Nolig, 4) = allActSur(Nolig)
    Next
    derligsheet = FL0NO.Cells.SpecialCells(xlCellTypeLastCell).Row
    For e = LigDeb0No To derligsheet
        FL0NO.Cells(e, EXTENDED).VALUE = ""
    Next
    For e = 1 To UBound(extendNb)
        FL0NO.Cells(LigDeb0No - 1 + newNumLig(e), EXTENDED).VALUE = extendNb(e)
    Next
    Dim entete() As String
    ReDim entete(1 To 1, 1 To 4)
    entete(1, 1) = "Feuille"
    entete(1, 2) = "Ligne"
    entete(1, 3) = "Entité"
    entete(1, 4) = ""
    
    FLNOM.Range("A1:D1").VALUE = entete
    FLNOM.Range("A2:c" & (UBound(allVecRet) + 1)).VALUE = aecrire
    sngChrono = Timer - sngChrono
    newLogLine = Array("NOMENCLATURE", "", "0NOM", "GENERATION", "FIN", time(), (Int(1000 * sngChrono) / 1000) & " s")
    logNom = alimLog(logNom, newLogLine)
    logNom = Application.Transpose(logNom)
    FLCNO.Range("A1:K" & UBound(logNom, 1)).VALUE = logNom
    derlig = Split(FLCNO.UsedRange.Address, "$")(4)
    Dim col As String
    col = "b"
    
    resColoriage = coloriage(LigDeb0No, FL0NO, FLCNO)
    
End Sub
Function getPre(entite As String, simpleStr() As String) As Boolean
    spl = Split(entite, ">")
    getPre = False
    achercher = ""
    If UBound(simpleStr) = 0 Then
        getPre = True
        Exit Function
    End If
    For I = LBound(spl) To UBound(spl)
        If I = LBound(spl) Then jct = ""
        If I <> LBound(spl) Then jct = ">"
        If I < UBound(spl) Then achercher = achercher & jct & spl(I)
    Next
    For s = LBound(simpleStr) To UBound(simpleStr)
        If Not simpleStr(s) Like "*.*" Then
            If achercher = "" Then
                getPre = True
                Exit For
            Else
                spla = Split(simpleStr(s), ":")
                If spla(0) = achercher Then
                    getPre = True
                    Exit For
                End If
            End If
        End If
    Next
End Function
Function testExtend(ent As String, ext() As String, num As Integer) As Boolean
    Dim newLogLine() As Variant
    testExtend = False
    Dim test() As Boolean
    spl = Split(ext(LBound(ext)), ".")
    ReDim test(LBound(spl) To UBound(spl))
    For I = LBound(spl) To UBound(spl)
        test(I) = False
    Next
'MsgBox Join(ext, " | ")
    For l = LBound(ext) To UBound(ext)
        spl = Split(ext(l), ".")
        For I = LBound(spl) To UBound(spl)
            If spl(I) = "" Then
                testExtend = True
                test(I) = True
'MsgBox i & ">>>" & ext(l)
            Else
                'test(i) = False
            End If
        Next
    Next
    If testExtend Then
        For I = LBound(test) To UBound(test)
            If test(I) Then
newLogLine = Array("", "ERREUR", "0NOM", "", "", time(), "", "THIS" & (1 + I), ent, num, "Facteur vide")
logNom = alimLog(logNom, newLogLine)
            End If
        Next
    End If
End Function
Function isInList(CH() As String, list() As String) As Boolean
    ' teste le AND de ch dans list
    isInList = True
    Dim itemInList As Boolean
    For c = LBound(CH) To UBound(CH)
        itemInList = False
        For I = LBound(list) To UBound(list)
            If list(I) = CH(c) Then
                itemInList = True
                Exit For
            End If
        Next
        isInList = isInList And itemInList
    Next
End Function
Function stringIsInList(str As String, list() As String) As Boolean
    stringIsInList = False
    For c = LBound(list) To UBound(list)
        If list(c) = str Then
            stringIsInList = True
            Exit For
        End If
    Next
End Function
Function getExtended(nomPerEquFeu As String, errorORalerte As Boolean, entite As String, simpleStr() As String, num As Integer, dimension As String, ORIGINE As String, colOrLig As String) As String()
' retourne les entités étendues de entite
    ' errorORalerte true pour
'If num = 12 And Left(entite, 1) = "B" Then MsgBox num & ":" & entite
    NUMENCOURS = num
'If num = 23 Then MsgBox nomPerEquFeu & Chr(10) & entite & Chr(10) & dimension
    Dim dimens As String
    dimens = Left(entite & ">", 2)
    If dimens = "a>" Or dimens = "s>" Then
        dimens = Left(dimens, 1)
    Else
        dimens = "n"
    End If
    Dim entiteInit As String
    entiteInit = entite
    Dim regEACH As VBScript_RegExp_55.RegExp
    Set regEACH = New VBScript_RegExp_55.RegExp
    Dim joinSpl() As String
    Dim res() As String
    ReDim res(1 To 1)
    If entite = "" Then
        res(1) = entite
        getExtended = res
        Exit Function
    End If
    If UBound(simpleStr) = 0 Then
        res(1) = entite
        getExtended = res
    Else
    '''Dim listFact(1 To 9) As Variant
    Dim listFact1() As String
    Dim listFact2() As String
    Dim listFact3() As String
    Dim listFact4() As String
    Dim listFact5() As String
    Dim listFact6() As String
    Dim listFact7() As String
    Dim listFact8() As String
    Dim listFact9() As String
    
    Dim cardFact(1 To 9) As Long
    For n = 1 To 9
        cardFact(n) = 0
    Next
    Dim lf() As String
    ReDim lf(0)
    Dim ind As Long
    Dim f As Integer
    Dim nbLig As Long
    nbLig = 1
    Dim resEntite As String
    resEntite = ""
    Dim okEntiteList() As Boolean
    Dim okEntite As Boolean
    where = ""
    Dim newLogLine() As Variant
    spl = Split(entite, ".")
    ReDim joinSpl(UBound(spl))
    For NoF = LBound(spl) To UBound(spl)
        joinSpl(NoF) = spl(NoF)
    Next
    f = 0
    Dim factAtt As String
    Dim itemAtt As String
    Dim fact As String
    colonne = "ATTR"
'If num = 64 Then MsgBox entite
    ReDim okEntiteList(LBound(spl) To UBound(spl))
'If num = 55 Then MsgBox num & ":" & entite & ">>>>>" & LBound(spl) & ":" & UBound(spl)
    For Each facteur In spl
        f = f + 1
        ReDim lf(1 To 1)
'If num = 55 Then MsgBox "deb:" & facteur
'If num = 55 Then MsgBox "facteur:" & entite & Chr(10) & facteur
        ' FIRST
'If num = 109 And nomPerEquFeu = "equ" Then MsgBox f & ":::" & entite & "=====" & facteur
        If facteur Like "*FIRST*" And InStr(facteur, ":FIRST(") = 0 Then
            factSpl = Split(facteur, "FIRST")
            regEACH.Pattern = factSpl(0)
            ind = 0
            For Each ligne In simpleStr
                If regEACH.test(ligne) And Not ligne Like "*.*" Then
                    If Not regEACH.Replace(ligne, "") Like "*>*" Then
                        factAtt = Split(facteur & ":", ":")(1)
'If num = 58 Then MsgBox entite & ":AVANT isAnAttribut=" & extraitAttribut(factAtt)
                        If dimens = "n" Then
                            aaa = isAnAttribut(extraitAttribut(factAtt), ATTRIBUTS, num, dimension, ORIGINE, colonne & f)
                        Else
                            aaa = True
                        End If
                        factCorps = Split(facteur & ":", ":")(0)
                        itemEach = regEACH.Replace(ligne, "")
                        itemAtt = Split(itemEach & ":", ":")(1)
'If num = 20 Then MsgBox entite & ":APRES isAnAttribut=" & aaa & "::" & extraitAttribut(factAtt) & ":::" & factAtt & ":::" & Join(ATTRIBUTS, "|")
                        resItem = Replace(facteur, "EACH", regEACH.Replace(ligne, ""))
                        If getAttInList(factAtt, itemAtt) Then
                            If ind > 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
                            lf(UBound(lf)) = Replace(factCorps, "FIRST", itemEach)
                            ind = ind + 1
                            Exit For
                        Else
                            If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                                If ind > 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
                                If factAtt <> "" Then factAtt = ":" & factAtt
                                lf(UBound(lf)) = Split(Replace(factCorps, "FIRST", itemEach), ":")(0) & factAtt
                                ind = ind + 1
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
            If ind = 0 Then
                If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                    okEntite = False
                    okEntiteList(f - 1) = okEntite
                Else
                    okEntite = True
                    okEntiteList(f - 1) = okEntite
                    If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                End If
            End If
        ' LAST
        ElseIf facteur Like "*LAST*" And InStr(facteur, ":LAST(") = 0 Then
            factSpl = Split(facteur, "LAST")
            regEACH.Pattern = factSpl(0)
            ind = 0
            For Each ligne In simpleStr
                If regEACH.test(ligne) And Not ligne Like "*.*" Then
                    If Not regEACH.Replace(ligne, "") Like "*>*" Then
                        factAtt = Split(facteur & ":", ":")(1)
                        If dimens = "n" Then
                            aaa = isAnAttribut(extraitAttribut(factAtt), ATTRIBUTS, num, dimension, ORIGINE, colonne & f)
                        Else
                            aaa = True
                        End If
                        factCorps = Split(facteur & ":", ":")(0)
                        itemEach = regEACH.Replace(ligne, "")
                        itemAtt = Split(itemEach & ":", ":")(1)
                        resItem = Replace(facteur, "LAST", regEACH.Replace(ligne, ""))
                        If getAttInList(factAtt, itemAtt) Then
                            'If ind = 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
                            lf(UBound(lf)) = Replace(factCorps, "LAST", itemEach)
                            ind = 1
                        End If
                    End If
                End If
            Next
            If ind = 0 Then
                If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                    okEntite = False
                    okEntiteList(f - 1) = okEntite
                Else
                    okEntite = True
                    okEntiteList(f - 1) = okEntite
                    If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                End If
            End If
        ' EACH
        ElseIf facteur Like "*EACH*" Then
            factSpl = Split(facteur, "EACH")
            regEACH.Pattern = factSpl(0)
            ind = 0
'If num = 125 Then MsgBox facteur & "==" & regEACH.Pattern & Chr(10) & Join(simpleStr, Chr(10))
            For Each ligne In simpleStr
                 
                If regEACH.test(ligne) And Not ligne Like "*.*" Then
                    If Not regEACH.Replace(ligne, "") Like "*>*" Then
                        factAtt = Split(facteur & ":", ":")(1)
                        If dimens = "n" Then
                            aaa = isAnAttribut(extraitAttribut(factAtt), ATTRIBUTS, num, dimension, ORIGINE, colonne & f)
                        Else
                            aaa = True
                        End If
'If num = 125 Then MsgBox facteur & "<<<>>>" & factAtt & "<*>" & extraitAttribut(factAtt) & "<<>>" & aaa
                        factCorps = Split(facteur & ":", ":")(0)
                        itemEach = regEACH.Replace(ligne, "")
                        itemAtt = Split(itemEach & ":", ":")(1)
                        resItem = Replace(facteur, "EACH", regEACH.Replace(ligne, ""))
'If num = 125 Then MsgBox ligne & Chr(10) & itemEach & Chr(10) & resItem
'MsgBox facteur & Chr(10) & ligne & Chr(10) & factAtt & ":" & itemAtt & ":" & getAttInList(factAtt, itemAtt)
                        If getAttInList(factAtt, itemAtt) Then
'If num = 125 Then MsgBox facteur & "<<<getAttInList>>>" & factAtt & "<*>" & extraitAttribut(factAtt) & "<<>>" & itemAtt
                            If ind > 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
'If num = 125 Then MsgBox factCorps & Chr(10) & itemEach
'MsgBox facteur & Chr(10) & ligne & Chr(10) & factAtt & ":" & itemAtt & ":" & getAttInList(factAtt, itemAtt) & Chr(10) & UBound(lf)
                            lf(UBound(lf)) = Replace(factCorps, "EACH", itemEach)
                            ind = ind + 1
                        End If
                    End If
                End If
            Next
            If ind = 0 Then
                If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                    okEntite = False
                    okEntiteList(f - 1) = okEntite
                Else
                    okEntite = True
                    okEntiteList(f - 1) = okEntite
                    If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                End If
            End If
        ' DESC
        ElseIf facteur Like "*DESC*" Then
            factSpl = Split(facteur, "DESC")
            regEACH.Pattern = "^" & factSpl(0)
            ind = 0
            For Each ligne In simpleStr
                If regEACH.test(ligne) And Not ligne Like "*.*" Then
                    If Not ligne Like "*.*" Then
                        factAtt = Split(facteur & ":", ":")(1)
                        If dimens = "n" Then
                            aaa = isAnAttribut(extraitAttribut(factAtt), ATTRIBUTS, num, dimension, ORIGINE, colonne & f)
                        Else
                            aaa = True
                        End If
                        factCorps = Split(facteur & ":", ":")(0)
                        itemEach = regEACH.Replace(ligne, "")
                        itemAtt = Split(itemEach & ":", ":")(1)
                        resItem = Replace(facteur, "DESC", regEACH.Replace(ligne, ""))
                        If getAttInList(factAtt, itemAtt) Then
                            If ind > 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
                            lf(UBound(lf)) = Replace(factCorps, "DESC", itemEach)
                            ind = ind + 1
                        End If
                    End If
                End If
            Next
'If num = 118 Then MsgBox "LEAF=" & ind & ":" & facteur & ":" & lf(UBound(lf))
            If ind = 0 Then
                If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                    okEntite = False
                    okEntiteList(f - 1) = okEntite
                Else
                    okEntite = True
                    okEntiteList(f - 1) = okEntite
                    If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                End If
            End If
        ' LEAF
        ElseIf facteur Like "*LEAF*" Then
            factSpl = Split(facteur, "LEAF")
            regEACH.Pattern = "^" & factSpl(0)
            ind = 0
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox num & ":" & facteur & ":" & UBound(simpleStr)
'If num = 55 Then MsgBox num & ":" & facteur & ":" & UBound(simpleStr)
            For Each ligne In simpleStr
                If regEACH.test(ligne) And Not ligne Like "*.*" Then
                    If Not ligne Like "*.*" Then
                        LigneSansAtt = Split(ligne, ":")(0)
                        aretenir = True
                        For Each ligne1 In simpleStr
                            If Not ligne1 Like "*.*" Then
                                Ligne1SansAtt = Split(ligne1, ":")(0)
'If num = 120 Then MsgBox LigneSansAtt & ":" & Len(LigneSansAtt) & "<<<" & ":" & Len(Ligne1SansAtt) & Ligne1SansAtt
                                If Len(LigneSansAtt) < Len(Ligne1SansAtt) Then
'If num = 120 Or num = 114 Then MsgBox num & ":" & Left(Ligne1SansAtt, Len(LigneSansAtt)) & "=" & LigneSansAtt
'''& ":" & Len(LigneSansAtt) & "<<<" & ":" & Len(Ligne1SansAtt) & Ligne1SansAtt
                                    If Left(Ligne1SansAtt, Len(LigneSansAtt)) = LigneSansAtt Then
                                        aretenir = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
'If num = 120 Or num = 114 Then MsgBox num & ":" & aretenir & ":" & LigneSansAtt
                        If aretenir Then
'If num = 120 Then MsgBox LigneSansAtt
                            factAtt = Split(facteur & ":", ":")(1)
'If num = 145 Then MsgBox "getExtended=" & ligne & ":::" & factAtt & ":::" & Join(ATTRIBUTS, "|")
                            If dimens = "n" Then
                                aaa = isAnAttribut(extraitAttribut(factAtt), ATTRIBUTS, num, dimension, ORIGINE, colonne & f)
                            Else
                                aaa = True
                            End If
                            factCorps = Split(facteur & ":", ":")(0)
                            itemEach = regEACH.Replace(ligne, "")
                            itemAtt = Split(itemEach & ":", ":")(1)
                            resItem = Replace(facteur, "LEAF", regEACH.Replace(ligne, ""))
'If num = 145 Then MsgBox "getExtended=" & ligne & ":" & aaa & "::" & extraitAttribut(factAtt) & ":::" & factAtt & ":::" & Join(ATTRIBUTS, "|")
                            If getAttInList(factAtt, itemAtt) Then
                                If ind > 0 Then ReDim Preserve lf(LBound(lf) To (UBound(lf) + 1))
                                lf(UBound(lf)) = Replace(factCorps, "LEAF", itemEach)
                                ind = ind + 1
                            End If
                        End If
'If num = 120 Then MsgBox LigneSansAtt
                    End If
                End If
            Next
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox num & ":" & ind
            If ind = 0 Then
                If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                    okEntite = False
                    okEntiteList(f - 1) = okEntite
                Else
                    okEntite = True
                    okEntiteList(f - 1) = okEntite
                    If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                End If
            End If
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox facteur & ":fin"
        Else
            fact = facteur
            primAtt = Split(facteur & ":", ":")(1)
'If num = 0 Then MsgBox entite & ":::" & facteur & ":::" & primAtt
            facteur = propagationAttributs(fact, simpleStr, num)
            lf(UBound(lf)) = facteur
            okEntite = True
            okEntiteList(f - 1) = okEntite
'If num = 20 And nomPerEquFeu = "equ" Then MsgBox fact & Chr(10) & facteur
            If dimension = "NOMENCLATURE" And UBound(spl) = 0 Then
                okEntite = False
                okEntiteList(f - 1) = okEntite
                If UBound(Split(facteur, ":")) > 0 Then
                    Dim att As String
                    att = Split(facteur, ":")(1)
'If num = 100 Then MsgBox att
                    If primAtt <> "" Then
                        If dimens = "n" Then
                            aaa = isAnAttribut(att, ATTRIBUTS, num, dimension, ORIGINE, colonne & "1")
                        Else
                            aaa = True
                        End If
                    End If
                End If
            Else
'If num = 0 Then MsgBox facteur & Chr(10) & UBound(simpleStr)
                For Each ligne In simpleStr
                    If Not ligne Like "*.*" Then
'If ligne = "" Then MsgBox num & ":" & dimension & ":" & origine
                        LigneSansAtt = Split(ligne, ":")(0)
                        facteurSansAtt = Split(facteur, ":")(0)
'If num = 0 Then MsgBox "getExtended:" & facteur & "===" & ligne & ">>>" & LigneSansAtt & ":" & facteurSansAtt
                        If facteurSansAtt = LigneSansAtt Then
'If num = 0 Then MsgBox "getExtended ok:" & facteur & ":" & LigneSansAtt & ":" & facteurSansAtt
                            splf = Split(facteur, ":")
                            If UBound(splf) = 0 Then
                                lf(UBound(lf)) = ligne
                                okEntite = False
                                okEntiteList(f - 1) = okEntite
                            Else
                                '''If UBound(Split(ligne, ":")) = 0 Then
                                    '''okEntite = True
                                    '''okEntiteList(f - 1) = okEntite
                                '''Else
                                    ' filtre attribut
                                    factAtt = Split(facteur & ":", ":")(1)
                                    itemEach = regEACH.Replace(ligne, "")
                                    itemAtt = Split(itemEach & ":", ":")(1)
                                    
                                    '''Dim s() As String
                                    '''Dim ss() As String
                                    '''Dim at As String
                                    '''at = Split(ligne, ":")(1)
                                    '''If primAtt <> "" Then
                                        '''If dimens = "n" Then
                                            '''aaa = isAnAttribut(at, ATTRIBUTS, num, dimension, origine, colonne & "1")
                                        '''Else
                                            '''aaa = True
                                        '''End If
                                    '''End If
                                    
                                    '''ss = Split(Split(ligne, ":")(1), ",")
                                    '''s = Split(splf(1), ",")
                                    '''a = isInList(s, ss)
'If num = 58 Then MsgBox "getExtended=====" & facteur & ">>>" & Join(s, ",") & "<>" & Join(ss, ",") & "<<<" & a
                                    If getAttInList(factAtt, itemAtt) Then
                                        lf(UBound(lf)) = ligne
                                        okEntite = False
                                        okEntiteList(f - 1) = okEntite
                                    Else
                                        okEntite = True
                                        okEntiteList(f - 1) = okEntite
                                    End If
'''If num = 20 Then MsgBox entite & Chr(10) & "getExtended=====" & facteur & ":" & LigneSansAtt & ":" & facteurSansAtt & ">>>" & okEntite
                                '''End If
                            End If
                        End If
                    End If
                Next
            End If
'If num = 0 Then MsgBox okEntite
            If okEntite Then
                If dimension = "NOMENCLATURE" Or dimension = "QUANTITE" Then where = where & "," & "BEFORE" & f
                If dimension = "AREA" Then where = "1AREA"
                If dimension = "SCENARIO" Then where = "1SCENARIO"
                If dimension = "EQUATION" Then where = "1EQUATION"
            End If
        End If
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin1:" & facteur
        '''''listFact(f) = lf
        If f = 1 Then listFact1 = lf
        If f = 2 Then listFact2 = lf
        If f = 3 Then listFact3 = lf
        If f = 4 Then listFact4 = lf
        If f = 5 Then listFact5 = lf
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin2:" & facteur
        If f = 6 Then listFact6 = lf
        If f = 7 Then listFact7 = lf
        If f = 8 Then listFact8 = lf
        If f = 9 Then listFact9 = lf
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin3:" & facteur
        cardFact(f) = UBound(lf)
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin4:" & nbLig
        nbLig = nbLig * UBound(lf)
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin5:" & facteur
'If num = 55 Then MsgBox "suite:" & facteur
'If num = 0 Then
'''If num = 20 And nomPerEquFeu = "equ" Then MsgBox f & Chr(10) & Join(listFact(f), Chr(10))
'MsgBox f & ">>>" & okEntite & ":" & facteur & Chr(10) & Join(lf, Chr(10))
'If num > 55 Then MsgBox num
'If num = 55 And Left(facteur, 6) = "USAGES" Then MsgBox f & ":" & UBound(lf) & ":fin:" & facteur
    Next
'If num = 55 Then MsgBox entite
'If num = 118 Then MsgBox ind & ":" & facteur & ":" & lf(UBound(lf))
    ReDim res(1 To nbLig)
    Dim nbL As Long
    nbL = 0
    Dim interm As String
    interm = ""
    For n1 = 1 To cardFact(1)
        interm = listFact1(n1)
        If cardFact(2) > 0 Then
            For n2 = 1 To cardFact(2)
interm = listFact1(n1) & "." & listFact2(n2)
                If cardFact(3) > 0 Then
                    For n3 = 1 To cardFact(3)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3)
                        If cardFact(4) > 0 Then
                            For n4 = 1 To cardFact(4)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4)
                                If cardFact(5) > 0 Then
                                    For n5 = 1 To cardFact(5)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4) & "." & listFact5(n5)
                                        If cardFact(6) > 0 Then
                                            For n6 = 1 To cardFact(6)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4) & "." & listFact5(n5) & "." & listFact6(n6)
                                                If cardFact(7) > 0 Then
                                                    For n7 = 1 To cardFact(7)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4) & "." & listFact5(n5) & "." & listFact6(n6) & "." & listFact7(n7)
                                                        If cardFact(8) > 0 Then
                                                            For n8 = 1 To cardFact(8)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4) & "." & listFact5(n5) & "." & listFact6(n6) & "." & listFact7(n7) & "." & listFact8(n8)
                                                                If cardFact(9) > 0 Then
                                                                    For n9 = 1 To cardFact(9)
interm = listFact1(n1) & "." & listFact2(n2) & "." & listFact3(n3) & "." & listFact4(n4) & "." & listFact5(n5) & "." & listFact6(n6) & "." & listFact7(n7) & "." & listFact8(n8) & "." & listFact9(n9)
                                                                        nbL = nbL + 1
                                                                        res(nbL) = interm
                                                                    Next
                                                                Else
                                                                    nbL = nbL + 1
                                                                    res(nbL) = interm
                                                                End If
                                                            Next
                                                        Else
                                                            nbL = nbL + 1
                                                            res(nbL) = interm
                                                        End If
                                                    Next
                                                Else
                                                    nbL = nbL + 1
                                                    res(nbL) = interm
                                                End If
                                            Next
                                        Else
                                            nbL = nbL + 1
                                            res(nbL) = interm
                                        End If
                                    Next
                                Else
                                    nbL = nbL + 1
                                    res(nbL) = interm
                                End If
                            Next
                        Else
                            nbL = nbL + 1
                            res(nbL) = interm
                        End If
                    Next
                Else
                    nbL = nbL + 1
                    res(nbL) = interm
                End If
            Next
        Else
            nbL = nbL + 1
            res(nbL) = interm
        End If
    Next
    getExtended = res
'If num = 20 And nomPerEquFeu = "equ" Then MsgBox entite & Chr(10) & "get1=" & Join(res, Chr(10))
    Dim resOkEntite As Boolean
    resOkEntite = False
    For b = LBound(okEntiteList) To UBound(okEntiteList)
        resOkEntite = resOkEntite Or okEntiteList(b)
    Next

    Dim newres() As String
    ReDim newres(0)
'MsgBox UBound(simpleStr)
    If UBound(spl) > 0 And errorORalerte And Not resOkEntite Then
    ' test si l'entité complexe est dans la nomenclature
        Dim nr As Integer
        n = 0
        For I = LBound(res) To UBound(res)
            For Each ligne In simpleStr
'If InStr(ligne, "SECTEUR>RESIDENTIEL>MENAGE>Maisons Individuelles.TECHNOLOGIES>") > 0 And InStr(res(i), "Maisons Individuelles") > 0 Then
    'MsgBox res(i) & Chr(10) & Chr(10) & ligne
'End If
                If res(I) = ligne Then
'MsgBox res(i) & Chr(10) & ligne
                    nr = nr + 1
                    If UBound(newres) = 0 Then ReDim newres(1 To 1)
                    If UBound(newres) >= 1 Then ReDim Preserve newres(1 To nr)
                    newres(UBound(newres)) = res(I)
'If num = 164 Then MsgBox nr & "=get=" & LBound(newres) & ":" & UBound(newres) & ":" & newres(UBound(newres)) & "<<<" & Join(newres, "|")
                    Exit For
                End If
            Next
        Next
'If num = 164 Then MsgBox "get TWO=" & UBound(newres) & ">>>" & Join(Res, "|")
'MsgBox "get2=" & Join(newres, Chr(10))
        If UBound(newres) = 0 Then
            resOkEntite = True
            If nomPerEquFeu = "equ" Or nomPerEquFeu = "feu" Then
                If nomPerEquFeu = "equ" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Equation", time(), "EQUATION", entiteInit, num, "terme de l'équation inconnu2") '
                End If
                If nomPerEquFeu = "feu" Then
                    If colOrLig = "L" Then ou = ""
                    If colOrLig = "C" Then ou = " C"
newLogLine = Array("", "ERREUR" & ou, ORIGINE, "Data", time(), "PERIMETRE", entiteInit, num, "Périmètre inconnu") '
                End If
            Else
If where = "AREA" Then
    If num > 0 Then newLogLine = Array("", "ERREUR", ORIGINE, "Extension AREA", time(), "", entite, num, "area inconnu")
ElseIf where = "SCENARIO" Then
    If num > 0 Then newLogLine = Array("", "ERREUR", ORIGINE, "Extension SCENARIO", time(), "", entite, num, "scénario inconnu")
Else
    If num > 0 Then newLogLine = Array("", "ERREUR", ORIGINE, "Extension NOMENCLATURE", time(), "", entite, num, "Entité inconnue")
End If
            End If
            If num > 0 Then logNom = alimLog(logNom, newLogLine)
            resOkEntite = False
            res = newres
            getExtended = res
        Else
            ReDim res(1 To UBound(newres))
            res = newres
            getExtended = res
'If num = 164 Then MsgBox "get TWO=" & Chr(10) & UBound(getExtended) & Chr(10) & Join(getExtended, Chr(10))
        End If
    End If
'If num = 164 Then MsgBox "get3=" & UBound(Res) & ":" & Join(Res, "|")
'MsgBox "get=" & Join(res, Chr(10))
'If num = 35 Then MsgBox entite & "===" & resOkEntite
    If UBound(res) = 1 Then
        If res(1) = "" Then
            okEntite = True
            'If dimension = "NOMENCLATURE" Then where = where & "," & "BEFORE" & f
            If dimension = "AREA" Then where = "1AREA"
            If dimension = "SCENARIO" Then where = "1SCENARIO"
            If dimension = "EQUATION" Then where = "1EQUATION"
        End If
    End If
    If resOkEntite Then
'If num = 164 Then MsgBox "get ICI " & Join(getExtended, Chr(10))
        If num = 0 Then
            ReDim getExtended(0)
        Else
        If UBound(joinSpl) > 0 And num > 0 Then
            ' cas produit cartésien multiple
            ReDim getExtended(0)
            where = Mid(where, 2)
'If num = 91 Then MsgBox num & ":" & "Facteur inconnu"
    If nomPerEquFeu = "bou" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Extension", time(), "BOUCLE", entite, num, "Modalité de boucle inconnue")
    Else
newLogLine = Array("", "ERREUR", ORIGINE, "Extension", time(), where, entite, num, "Facteur inconnu")
    End If
logNom = alimLog(logNom, newLogLine)
        Else
            ' cas entité simple
'If num = 8 Then MsgBox "getExtended:" & errorORalerte & where
            If errorORalerte And num > 0 Then
                ReDim getExtended(0)
                where = Mid(where, 2)
                If nomPerEquFeu = "equ" Or nomPerEquFeu = "feu" Then
                    If nomPerEquFeu = "equ" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Equation", time(), "EQUATION", entiteInit, num, "terme de l'équation inconnu3")
                    End If
                    If nomPerEquFeu = "feu" Then
                    If colOrLig = "L" Then ou = ""
                    If colOrLig = "C" Then ou = " C"
newLogLine = Array("", "ERREUR" & ou, ORIGINE, "Data", time(), "PERIMETRE", entiteInit, num, "Périmètre inconnu") '
                    End If
                Else
If where = "AREA" Then
    newLogLine = Array("", "ERREUR", ORIGINE, "Extension AREA", time(), where, entite, num, "area inconnu")
ElseIf where = "SCENARIO" Then
    newLogLine = Array("", "ERREUR", ORIGINE, "Extension SCENARIO", time(), where, entite, num, "scénario inconnu")
Else
    If nomPerEquFeu = "bou" Then
newLogLine = Array("", "ERREUR", ORIGINE, "Extension", time(), "BOUCLE", entite, num, "Modalité de boucle inconnue")
    Else
newLogLine = Array("", "ERREUR", ORIGINE, "Extension NOMENCLATURE", time(), where, entite, num, "Entité inconnue2")
    End If
End If
                End If
                logNom = alimLog(logNom, newLogLine)
            End If
        End If
        End If
    End If
    '''If errorORalerte Then
        
                '''ReDim getExtended(0)
                '''where = Mid(where, 2)
'''newLogLine = Array("", "ERREUR", origine, "Extension", Time(), where, entite, num, "Entité inconnue")
'''logNom = alimLog(logNom, newLogLine)
    '''End If
    End If
    If num = 0 And resOkEntite Then
        ReDim getExtended(0)
    End If
'If num = 0 Then MsgBox resOkEntite & " FIN getExtended=" & UBound(getExtended)
End Function
Function getAttInList(att As String, list As String) As Boolean
    natt = extraitAttribut(att)
    'If att = "Gaz" Then MsgBox natt
    typeAtt = "AND"
    If InStr(att, "NOT(") = 1 Then typeAtt = "NOT"
    If InStr(att, "AND(") = 1 Then typeAtt = "AND"
    If InStr(att, "OR(") = 1 Then typeAtt = "OR"
    If InStr(att, "FIRST(") = 1 Then typeAtt = "FIRST"
    If InStr(att, "LAST(") = 1 Then typeAtt = "LAST"
    boolEnCours = True
    If att = "" Then
        getAttInList = True
    Else
        ' cas du FIRST
        If typeAtt = "FIRST" Then
            spla = Split(natt & ",", ",")
            spll = Split(list, ",")
            getAttInList = False
            a = spla(0)
            boolEnCours = False
            For Each l In spll
                If l = a Then
                    boolEnCours = True
                    Exit For
                End If
            Next
            getAttInList = boolEnCours
        End If
        ' cas du LAST
        If typeAtt = "LAST" Then
            spla = Split(natt & ",", ",")
            spll = Split(list, ",")
            getAttInList = False
            a = spla(UBound(spla) - 1)
            boolEnCours = False
            For Each l In spll
                If l = a Then
                    boolEnCours = True
                    Exit For
                End If
            Next
            getAttInList = boolEnCours
        End If
        ' cas du OR
        If typeAtt = "OR" Then
            spla = Split(natt, ",")
            spll = Split(list, ",")
            getAttInList = False
            For Each a In spla
                boolEnCours = False
                For Each l In spll
                    If l = a Then
                        boolEnCours = True
                        Exit For
                    End If
                Next
                getAttInList = getAttInList Or boolEnCours
            Next
'MsgBox att & "<>" & list & "<>" & getAttInList
        End If
        ' cas du AND (défaut
        If typeAtt = "AND" Then
            spla = Split(natt, ",")
            spll = Split(list, ",")
            getAttInList = True
            For Each a In spla
                boolEnCours = False
                For Each l In spll
                    If l = a Then
                        boolEnCours = True
                        Exit For
                    End If
                Next
                getAttInList = getAttInList And boolEnCours
            Next
        End If
        ' cas du NOT
        If typeAtt = "NOT" Then
            spla = Split(natt, ",")
            spll = Split(list, ",")
            getAttInList = False
'MsgBox ">" & natt & Chr(10) & list & "<"
            For Each a In spla
                boolEnCours = True
                For Each l In spll
'MsgBox att & ">" & natt & Chr(10) & list & "<" & l & ":" & a
                    If l = a Then
                        boolEnCours = False
'MsgBox att & ">" & natt & Chr(10) & list & "<" & l & ":" & a
                        Exit For
                    End If
                Next
                getAttInList = getAttInList Or boolEnCours
            Next
        End If
    End If
'If NUMENCOURS = 123 Then MsgBox att & ":::" & natt & ":::" & getAttInList & ":::" & list
End Function
Sub SetNomenclatureOLD()
    Etape = 2
    Call DebutEtape(Etape)
    ''''On Error GoTo ErrorHandler
    Call SetFl              ' Initialisation des feuilles
    Call DelNomenclature    ' Initialisation de la feuille NOMENCLATURE
    derlig = Split(FL0NO.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FL0NO.Cells(Nolig, 1) = "*" Then
            LigDeb0No = Nolig + 1
            Exit For
        End If
    Next
    derlig = Split(FLNOM.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLNOM.Cells(Nolig, 1) = "*" Then
            LigDebNom = Nolig + 1
            Exit For
        End If
    Next
    'Alimentation des tableaux des classes et autres
    Dim Cla() As Variant
    With FL0NO.Range("a1").CurrentRegion
        ReDim Cla(.Rows.Count, .Columns.Count)
        Cla = .VALUE
    End With
    Call EcritureInput(Etape, "" & (UBound(Cla, 1) - LigDeb0No + 1))
    'Construction des classes
    ReDim CLASSES(1 To 1)
    ReDim CLASSESI(1 To 1)
    ReDim CLASSESN(1 To 1)
    Dim positions() As Integer
    ReDim positions(1 To UBound(Cla, 1))
    Dim NbEtendu() As Integer
    ReDim NbEtendu(1 To UBound(Cla, 1))
    Dim ListItems() As String
    Dim LastClasse As String
    LastClasse = ""
    Dim ClasseEnCours As String
    ClasseEnCours = ""
    Dim ItemEnCours As String
    ItemEnCours = ""
    Dim NbClasse As Integer
    NbClasse = 0
    Dim NbItem As Integer
    NbItem = 0
    Dim NbLignes As Integer
    NbLignes = 0
    Dim NoLigne As Integer
    NoLigne = 0
    Dim ClasseToSearch As String
    Dim Go As Boolean
    Go = False
    For Nolig = LBound(Cla) To UBound(Cla)
        positions(Nolig) = NoLigne
        If Cla(Nolig, 1) = "*" Then
            Go = True
            positions(Nolig) = NbLignes
            NbEtendu(Nolig) = 0
        Else
            If Go Then
                If Left(Cla(Nolig, 1), 1) <> ":" Then
                    If UBound(Split(Cla(Nolig, 1), ">")) > 0 And Cla(Nolig, 2) = "" Then
                        ClasseEnCours = Split(Cla(Nolig, 1), ">")(0)
                        ItemEnCours = Split(Cla(Nolig, 1), ">")(1)
                        If ClasseEnCours <> LastClasse Then
                            NbClasse = NbClasse + 1
                            ReDim Preserve CLASSES(LBound(CLASSES) To NbClasse)
                            ReDim Preserve CLASSESI(LBound(CLASSESI) To NbClasse)
                            ReDim Preserve CLASSESN(LBound(CLASSESN) To NbClasse)
                            CLASSES(UBound(CLASSES)) = ClasseEnCours
                            NbItem = 0
                        End If
                        NbItem = NbItem + 1
                        ReDim Preserve ListItems(1 To NbItem)
                        ListItems(NbItem) = ItemEnCours
                        CLASSESI(UBound(CLASSESI)) = ListItems
                        CLASSESN(UBound(CLASSESN)) = UBound(ListItems)
                        LastClasse = ClasseEnCours
                        NbLignes = NbLignes + 1
                        positions(Nolig) = NbLignes
                        NbEtendu(Nolig) = 1
                    End If
                    If UBound(Split(Cla(Nolig, 1), ">")) < 1 Then
                        NbLignes = NbLignes + 1
                        positions(Nolig) = NbLignes
                        NbEtendu(Nolig) = 1
                    End If
                    If UBound(Split(Cla(Nolig, 1), ">")) > 0 And Cla(Nolig, 2) <> "" Then
                        ClasseToSearch = Split(Cla(Nolig, 1), ">")(0)
                        NbItem = 1
                        For NoCol = 1 To 3
                            If UBound(Split(Cla(Nolig, NoCol), ">")) > 0 Then
                                If Split(Cla(Nolig, NoCol), ">")(1) = "EACH" Then
                                    ClasseToSearch = Split(Cla(Nolig, NoCol), ">")(0)
                                    NoItems = GetPosItemInList(ClasseToSearch, CLASSES)
                                    If NoItems = -1 Then
                                        MsgBox "La classe " & ClasseToSearch & " n'existe pas"
                                    Else
                                        NbItem = NbItem * UBound(CLASSESI(NoItems))
                                    End If
                                End If
                            End If
                        Next
                        positions(Nolig) = NbLignes + 1
                        NbEtendu(Nolig) = NbItem
                        NbLignes = NbLignes + NbItem
                    End If
                Else
                    NbLignes = NbLignes + 1
                    positions(Nolig) = NbLignes
                    NbEtendu(Nolig) = 1
                End If
            End If
        End If
    Next
    Dim ResTab() As String
    ReDim ResTab(1 To NbLignes, 1 To 4)
    For Nolig = LBound(positions) To UBound(positions)
        If positions(Nolig) <> 0 Then
            ResTab(positions(Nolig), 1) = Cla(Nolig, 1)
            ResTab(positions(Nolig), 2) = Cla(Nolig, 2)
            ResTab(positions(Nolig), 3) = Cla(Nolig, 3)
            ResTab(positions(Nolig), 4) = NbEtendu(Nolig)
            FL0NO.Cells(Nolig, 4).VALUE = NbEtendu(Nolig)
        End If
    Next
    Dim TabInt() As String
    Dim listEach() As String
    Dim ReplaceItem As String
    Dim Pas As Integer
    Pas = 1
    For Nolig = 1 To NbLignes
        concat = ResTab(Nolig, 1) & "." & ResTab(Nolig, 2) & "." & ResTab(Nolig, 3)
        If InStr(concat, "EACH") > 0 Then
            NbEt = ResTab(Nolig, 4)
            ReDim listEach(1 To 3)
            listEach(1) = ResTab(Nolig, 1)
            listEach(2) = ResTab(Nolig, 2)
            listEach(3) = ResTab(Nolig, 3)
            Pas = 1
            For NoCol = 3 To 1 Step -1
                If UBound(Split(listEach(NoCol), ">")) > 0 Then
                    If Split(listEach(NoCol), ">")(1) = "EACH" Then
                        ClasseToSearch = Split(listEach(NoCol), ">")(0)
                        NoItems = GetPosItemInList(ClasseToSearch, CLASSES)
                        If NoItems = -1 Then
                            MsgBox "La classe " & ClasseToSearch & " n'existe pas"
                        Else
                            ReplaceItem = ResTab(Nolig, NoCol)
                            NbIt = UBound(CLASSESI(NoItems))
                            enPLus = 0
                            For NoC = 1 To Int(NbEt / (NbIt * Pas))
                                For NoIt = 1 To NbIt
                                    Replaced = Replace(ReplaceItem, "EACH", CLASSESI(NoItems)(NoIt))
                                    For NoPas = 1 To Pas
                                        LigOutput = Nolig + enPLus
                                        ResTab(LigOutput, NoCol) = Replaced
                                        enPLus = enPLus + 1
                                    Next
                                Next
                            Next
                            Pas = Pas * UBound(CLASSESI(NoItems))
                        End If
                    End If
                End If
            Next
        End If
    Next
    For Nolig = 1 To NbLignes
        If ResTab(Nolig, 2) = "" Then ResTab(Nolig, 4) = ResTab(Nolig, 1)
        If ResTab(Nolig, 2) <> "" Then ResTab(Nolig, 4) = ResTab(Nolig, 1) & "." & ResTab(Nolig, 2)
        If ResTab(Nolig, 3) <> "" Then ResTab(Nolig, 4) = ResTab(Nolig, 4) & "." & ResTab(Nolig, 2)
    Next
    FLNOM.Range("A" & LigDebNom & ":D" & NbLignes + LigDebNom - 1).VALUE = ResTab
    'NOMENCLATURE(NoLig - LigDebNom) = FLNOM.Cells(NoLig, 4)
    NbFin = NbLignes
    Call EcritureResultats(Etape, "NOMENCLATURE", "" & NbFin)
    Exit Sub
errorHandler: Call ErrorToDo(Etape, "NOMENCLATURE", "" & NbFin, Err)
End Sub
Function GetFirstLine(FL As Worksheet) As Integer()
'Renvoie la 1ère et dernière ligne de FL
    Dim FirstLine As Integer
    Dim LastLine As Integer
    Dim res() As Integer
    ReDim res(0 To 1)
    ReDim GetFirstLine(0 To 1)
    FirstLine = 1
    LastLine = Split(FL.UsedRange.Address, "$")(4)
    For Nolig = 1 To LastLine
        If FL.Cells(Nolig, 1) = "*" Then
            FirstLine = Nolig + 1
        End If
        If Left(FL.Cells(Nolig, 1), 3) = "END" Then
            LastLine = Nolig
            Exit For
        End If
    Next
    res(0) = FirstLine
    res(1) = LastLine
    GetFirstLine = res()
End Function
Function GetFirstCol(FL As Worksheet) As Integer()
'Renvoie la 1ère et dernière colonne de FL
    Dim FirstCol As Integer
    Dim LastCol As Integer
    Dim res() As Integer
    ReDim res(0 To 1)
    ReDim GetFirstCol(0 To 1)
    FirstCol = 1
    LastCol = Columns(Split(FL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To LastCol
        If Left(FL.Cells(1, NoCol).VALUE, 3) = "END" Then
            LastCol = NoCol
            Exit For
        End If
    Next
    res(0) = FirstCol
    res(1) = LastCol
    GetFirstCol = res()
End Function
Function GetClasseGauche(sep As String, chaine As String) As String
    spl = Split(chaine, ">" & sep)
    If UBound(spl) > 0 Then
        ChaineRev = StrReverse(spl(0))
        PosPoint = InStr(ChaineRev, ".")
        PosCroch = InStr(ChaineRev, "[")
        PosDeuxp = InStr(ChaineRev, ":")
        If PosPoint = 0 And PosCroch <> 0 Then SepFinal = "["
        If PosCroch = 0 And PosPoint <> 0 Then SepFinal = "."
        If PosCroch <> 0 And PosPoint <> 0 Then
            If PosPoint < PosCroch Then SepFinal = "."
            If PosCroch < PosPoint Then SepFinal = "["
        End If
        If PosDeuxp <> 0 Then SepFinal = ":"
        ChaineFinale = Split(ChaineRev, SepFinal)(0)
        GetClasseGauche = StrReverse(ChaineFinale)
    Else
        GetClasseGauche = ""
    End If
End Function
Function SetThis(Formule As String, listS As String) As String
    'Cla = GetClasseGauche("THIS", Formule)
    Dim list() As String
    list = Split(listS, ".")
    SetThis = Formule
    For No = LBound(list) To UBound(list)
        SetThis = Replace(SetThis, Split(list(No), ">")(0) & ">THIS", list(No))
    Next
End Function
Function SetAll(Formule As String) As String
'Remplace les ALL par les entités associées (ne marche que pour un seul ALL)
    SetAll = Formule
    If InStr(Formule, "[") <> 0 Then
        Gauche = Left(Formule, InStr(Formule, "[") - 1)
        PosDroite = Len(Formule) - InStr(Formule, ")")
        chaine = Mid(Formule, InStr(Formule, "["), InStr(Formule, ")") - InStr(Formule, "[") + 1)
        If PosDroite > -1 Then
            Droite = Right(Formule, PosDroite)
        End If
    End If
    Dim SplAll() As String
    SplAll = Split(chaine, ">ALL")
    If UBound(SplAll) > 0 Then
        Dim Cl As String
        SplAll2 = Split(SplAll(0), ">")
        If UBound(SplAll2) > 0 Then
            Cl = SplAll2(1)
        Else
            Cl = SplAll(0)
        End If
        SplAll3 = Split(Cl, ".")
        If UBound(SplAll3) > 0 Then
            Cl = SplAll3(1)
        End If
        If Left(Cl, 1) = "[" Then Cl = Mid(Cl, 2)
        ListChildren = GetChildrenClasse(Cl)
        Dim NewFormule As String
        NewFormule = ""
        For Each elt In ListChildren
            NewFormule = NewFormule & ";" & Replace(chaine, "ALL", elt)
        Next
        NewFormule = Mid(NewFormule, 2)
        SetAll = Gauche & NewFormule & Droite
    End If
End Function
Sub SetEquations()
    Etape = 3
    Dim NbFin As Integer
    Call DebutEtape(Etape)
    On Error GoTo errorHandler
    Call SetFl              ' Initialisation des feuilles
    Call DelEquations       ' Initialisation de la feuille EQUATIONS
    Dim derlig As Integer
    Dim dercol As Integer
    Dim LigDebEqu As Integer
    Dim LigDeb0Eq As Integer
    LigDebEqu = GetFirstLine(FLEQU)(0)
    LigDeb0Eq = GetFirstLine(FL0EQ)(0)
    derlig = Split(FL0EQ.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FL0EQ.UsedRange.Address, "$")(3)).Column
    'Alimentation des tableaux des équations génériques
    Dim Cla() As Variant
    ReDim Cla(1 To derlig - LigDeb0Eq + 1, 1 To dercol)
    Cla = FL0EQ.Range("A" & LigDeb0Eq & ":" & DecAlph(dercol) & derlig).VALUE
    Call EcritureInput(Etape, "" & UBound(Cla, 1))
    
    'Construction du tableau des équations étendues
    'ReDim CLASSESI(1 To 1)
    'ReDim CLASSESN(1 To 1)
    Dim positions() As Integer
    ReDim positions(1 To UBound(Cla, 1))
    Dim NbEtendu() As Integer
    ReDim NbEtendu(1 To UBound(Cla, 1))
    Dim ListItems() As String
    Dim LastClasse As String
    LastClasse = ""
    Dim ClasseEnCours As String
    ClasseEnCours = ""
    Dim ItemEnCours As String
    ItemEnCours = ""
    Dim NbClasse As Integer
    NbClasse = 0
    Dim NbItem As Integer
    NbItem = 0
    Dim NbLignes As Integer
    NbLignes = 0
    Dim NoLigne As Integer
    NoLigne = 0
    Dim ClasseToSearch As String
    For Nolig = LBound(Cla) To UBound(Cla)
        positions(Nolig) = NoLigne
        If Left(Cla(Nolig, 1), 1) <> ":" Then
            If UBound(Split(Cla(Nolig, 1), ">")) < 1 Then
                NbLignes = NbLignes + 1
                positions(Nolig) = NbLignes
                NbEtendu(Nolig) = 1
            End If
            If UBound(Split(Cla(Nolig, 1), ">")) > 0 Then
                ClasseToSearch = Split(Cla(Nolig, 1), ">")(0)
                NbItem = 1
                For NoCol = 1 To 3
                    If UBound(Split(Cla(Nolig, NoCol), ">")) > 0 Then
                        If Split(Cla(Nolig, NoCol), ">")(1) = "EACH" Then
                            ClasseToSearch = Split(Cla(Nolig, NoCol), ">")(0)
                            NoItems = GetPosItemInList(ClasseToSearch, CLASSES)
                            If NoItems = -1 Then
                                MsgBox "La classe " & ClasseToSearch & " n'existe pas"
                            Else
                                NbItem = NbItem * UBound(CLASSESI(NoItems))
                            End If
                        End If
                    End If
                Next
                positions(Nolig) = NbLignes + 1
                NbEtendu(Nolig) = NbItem
                NbLignes = NbLignes + NbItem
            End If
        Else
            NbLignes = NbLignes + 1
            positions(Nolig) = NbLignes
            NbEtendu(Nolig) = 1
        End If
    Next

   
    Dim nom As String
    Dim Unite As String
    Dim Echelle As String
    Dim Formule As String
    Dim FeuilleHisto As String
    Dim FormuleHisto As String
    Dim FeuilleJalon As String
    Dim FormuleJalon As String
    Dim ResTab() As String
    ReDim ResTab(1 To NbLignes, 1 To 12)
    For Nolig = LBound(positions) To UBound(positions)
        If positions(Nolig) <> 0 Then
            ResTab(positions(Nolig), 1) = Cla(Nolig, 1)
            ResTab(positions(Nolig), 2) = Cla(Nolig, 2)
            ResTab(positions(Nolig), 3) = Cla(Nolig, 3)
            ResTab(positions(Nolig), 4) = NbEtendu(Nolig)
            ''''FL0EQ.Cells(NoLig, 4).Value = NbEtendu(NoLig)
            ResTab(positions(Nolig), 5) = Cla(Nolig, 5)
            ResTab(positions(Nolig), 6) = Cla(Nolig, 6)
            ResTab(positions(Nolig), 7) = Cla(Nolig, 7)
            ResTab(positions(Nolig), 8) = Cla(Nolig, 8)
            ResTab(positions(Nolig), 9) = Cla(Nolig, 9)
            ResTab(positions(Nolig), 10) = Cla(Nolig, 10)
            ResTab(positions(Nolig), 11) = Cla(Nolig, 11)
            ResTab(positions(Nolig), 12) = Cla(Nolig, 12)
        End If
    Next
    Dim TabInt() As String
    Dim listEach() As String
    Dim ReplaceItem As String
    Dim Pas As Integer
    Pas = 1
    For Nolig = 1 To NbLignes
        concat = ResTab(Nolig, 1) & "." & ResTab(Nolig, 2) & "." & ResTab(Nolig, 3)
        If Left(ResTab(Nolig, 1), 1) <> ":" Then
        If InStr(concat, "EACH") > 0 Then
            NbEt = ResTab(Nolig, 4)
            ReDim listEach(1 To 3)
            listEach(1) = ResTab(Nolig, 1)
            listEach(2) = ResTab(Nolig, 2)
            listEach(3) = ResTab(Nolig, 3)
            Pas = 1
            nom = ResTab(Nolig, 5)
            Unite = ResTab(Nolig, 6)
            Echelle = ResTab(Nolig, 7)
            Formule = ResTab(Nolig, 8)
            FormuleBefore = ResTab(Nolig, 8)
            FeuilleHisto = ResTab(Nolig, 9)
            FormuleHisto = ResTab(Nolig, 10)
            FeuilleJalon = ResTab(Nolig, 11)
            FormuleJalon = ResTab(Nolig, 12)
            For NoCol = 3 To 1 Step -1
                If UBound(Split(listEach(NoCol), ">")) > 0 Then
                    If Split(listEach(NoCol), ">")(1) = "EACH" Then
                        ClasseToSearch = Split(listEach(NoCol), ">")(0)
                        NoItems = GetPosItemInList(ClasseToSearch, CLASSES)
                        If NoItems = -1 Then
                            MsgBox "La classe " & ClasseToSearch & " n'existe pas"
                        Else
                            ReplaceItem = ResTab(Nolig, NoCol)
                            NbIt = UBound(CLASSESI(NoItems))
                            enPLus = 0
                            For NoC = 1 To Int(NbEt / (NbIt * Pas))
                                For NoIt = 1 To NbIt
                                    Replaced = Replace(ReplaceItem, "EACH", CLASSESI(NoItems)(NoIt))
                                    'Formule = FormuleBefore
                                    For NoPas = 1 To Pas
                                        LigOutput = Nolig + enPLus
                                        ResTab(LigOutput, NoCol) = Replaced
                                        ResTab(LigOutput, 5) = nom
                                        ResTab(LigOutput, 6) = Unite
                                        ResTab(LigOutput, 7) = Echelle
'If InStr(Formule, ">THIS") > 0 And InStr(Formule, "evolutionTotal") > 0 Then MsgBox ClasseToSearch & ">" & CLASSESI(NoItems)(NoIt) & " DANS " & Formule
                                        'Formule = Replace(Formule, ClasseToSearch & ">THIS", ClasseToSearch & ">" & CLASSESI(NoItems)(NoIt))
                    
                                        Formule = SetAll(Formule)
                                        ResTab(LigOutput, 8) = Formule
                                        ResTab(LigOutput, 9) = FeuilleHisto
                                        'FormuleHisto = Replace(FormuleHisto, ClasseToSearch & ">THIS", ClasseToSearch & ">" & CLASSESI(NoItems)(NoIt))
                                        ResTab(LigOutput, 10) = FormuleHisto
                                        ResTab(LigOutput, 11) = FeuilleJalon
                                        'FormuleJalon = Replace(FormuleJalon, ClasseToSearch & ">THIS", ClasseToSearch & ">" & CLASSESI(NoItems)(NoIt))
                                        ResTab(LigOutput, 12) = FormuleJalon
                                        enPLus = enPLus + 1
                                    Next
                                Next
                            Next
                            Pas = Pas * UBound(CLASSESI(NoItems))
                        End If
                    End If
                End If
            Next
        End If
        End If
    Next
    ReDim EQUATIONS(1 To NbLignes)
    ReDim EQUATIONSQ(1 To NbLignes)
    ReDim EQUATIONSF(1 To NbLignes)
    For Nolig = 1 To NbLignes
        If ResTab(Nolig, 2) = "" Then ResTab(Nolig, 4) = ResTab(Nolig, 1)
        If ResTab(Nolig, 2) <> "" Then ResTab(Nolig, 4) = ResTab(Nolig, 1) & "." & ResTab(Nolig, 2)
        If ResTab(Nolig, 3) <> "" Then ResTab(Nolig, 4) = ResTab(Nolig, 4) & "." & ResTab(Nolig, 2)
        ResTab(Nolig, 8) = SetThis(ResTab(Nolig, 8), ResTab(Nolig, 4))
        ResTab(Nolig, 10) = SetThis(ResTab(Nolig, 10), ResTab(Nolig, 4))
        ResTab(Nolig, 10) = Replace(ResTab(Nolig, 10), "REF", ResTab(Nolig, 4))
        ResTab(Nolig, 12) = SetThis(ResTab(Nolig, 12), ResTab(Nolig, 4))
        ResTab(Nolig, 12) = Replace(ResTab(Nolig, 12), "REF", ResTab(Nolig, 4))
        EQUATIONS(Nolig) = ResTab(Nolig, 4)
        EQUATIONSQ(Nolig) = ResTab(Nolig, 5)
        EQUATIONSF(Nolig) = ResTab(Nolig, 8)
    Next
    
    FLEQU.Range("A" & LigDebEqu & ":" & DecAlph(UBound(ResTab, 2)) & NbLignes + LigDebEqu - 1).VALUE = ResTab
   

   
    NbFin = UBound(ResTab, 1)
    Call EcritureResultats(Etape, "EQUATIONS", "" & NbFin)
    Exit Sub
errorHandler: Call ErrorToDo(Etape, "EQUATIONS", "" & NbFin, Err)
End Sub
Function Instanciation(Item As Variant, Chaque As String, ThisToReplace As String) As Variant
'   Attention remplacer plutot Class>THIS que THIS car il peut y en avoir d'autres
    For cmpt1 = LBound(Item, 1) To UBound(Item, 1)
        For cmpt2 = LBound(Item, 2) To UBound(Item, 2)
            Item(cmpt1, cmpt2) = Replace(Item(cmpt1, cmpt2), ThisToReplace, Chaque)
        Next
    Next
    Instanciation = Item
End Function
Function GetEndLoop(FL As Variant, deb As Integer) As Integer
'   Retourne la ligne de la fin de boucle même si boucles imbriquées
    NbEndTo = 0
    NbEndOn = 0
    EndLoop = 0
    derlig = UBound(FL, 1)
    fin = UBound(FL, 2)
    For NoLigEnd = deb To derlig
        var = FL(NoLigEnd, fin)
        If InStr(var, "(") > 0 Then NbEndTo = NbEndTo + 1
        If UBound(Split(var, ")")) > 0 Then
            NbEndOn = NbEndOn + UBound(Split(var, ")"))
            If NbEndOn = NbEndTo Then
                EndLoop = NoLigEnd
                Exit For
            End If
        End If
    Next
    GetEndLoop = EndLoop
End Function
Function DelDebEndLoop(FL As Variant) As Variant
'   Elimine les id de la boucle englobante
    cmpt1 = UBound(FL, 1)
    NbEndTo = 0
    NbEndOn = 0
    EndLoop = 0
    Dim Inverse As String
    For cmpt2 = LBound(FL, 2) To UBound(FL, 2)
        var = FL(cmpt1, cmpt2)
        NbEndTo = NbEndTo + UBound(Split(var, "("))
        NbEndOn = NbEndOn + UBound(Split(var, ")"))
        If NbEndOn = NbEndTo Then
            EndLoop = cmpt2
            Inverse = StrReverse(FL(cmpt1, cmpt2))
            pos = InStr(Inverse, ")")
            PosToDel = Len(Inverse) - pos
FL(cmpt1, cmpt2) = Left(FL(cmpt1, cmpt2), PosToDel) & Right(FL(cmpt1, cmpt2), Len(FL(cmpt1, cmpt2)) - PosToDel - 1)
            Exit For
        End If
    Next
    SplitNum = Split(FL(cmpt1, 1), "°")
    If UBound(SplitNum) > 0 Then
        FL(cmpt1, 1) = "°" & SplitNum(1)
    Else
        FL(cmpt1, 1) = ""
    End If
    DelDebEndLoop = FL
End Function
Function GetItem(chaine As String) As String
    Dim res As String
    res = Right(chaine, Len(chaine) - 3)
    fin = InStr(res, "°")
    If fin <> 0 Then res = Mid(res, 1, fin - 1)
    fin = InStr(res, ")")
    If fin <> 0 Then res = Mid(res, 1, fin - 1)
    GetItem = res
End Function
Function Deplier(FL As Variant) As Variant
'   Déplie un niveau de boucle du tableau FL transposé pour pouvoir le redimensionner avec ReDim
    Dim DebLoop As Integer
    Dim NoBoucle As String
    derlig = UBound(FL, 2)
    ENDSSK = UBound(FL, 1)
    Dim ItemLoop() As Variant
    Dim ItemLoopOn() As Variant
    Dim FLOUT() As Variant
    ReDim FLOUT(1 To UBound(FL, 1), 1 To 1)
    Dim Formats() As Integer
    ReDim Formats(1 To 1)
    For Nolig = LBound(FL, 2) To UBound(FL, 2)
        var = FL(ENDSSK, Nolig)
        If Left(var, 1) = "(" Then
            NoBoucle = Right(Left(var, 2), 1)
            DebLoop = Nolig
            Dim Result() As String
            Dim ext As String
            Dim Arg As String
            Arg = var
            ext = GetItem(Arg)
            Result() = GetChildren(ext)
            'Détermination de la fin de la boucle
            Dim NbEnd As Integer
            EndLoop = GetEndLoop(Application.Transpose(FL), DebLoop)
            If EndLoop = 0 Then
                MsgBox ("Boucle sans )")
                Exit For
            End If
            ReDim ItemLoop(1 To UBound(FL, 1), 1 To EndLoop - Nolig + 1)
            For co = LBound(FL, 1) To UBound(FL, 1)
                For No = Nolig To EndLoop
                    ItemLoop(co, No - Nolig + 1) = FL(co, No)
                Next
            Next
            For NoCha = LBound(Result) To UBound(Result)
                ItemLoopOn = ItemLoop
                ItemLoopOnDel = DelDebEndLoop(ItemLoopOn)
                ItemLoopOn = ItemLoopOnDel
                Dim ItemLoopIns() As Variant
                ItemLoopIns = Instanciation(ItemLoopOn, CStr(Result(NoCha)), "THIS" & NoBoucle)
                ColIter = 0
                LigIter = 0
                NewLig = UBound(FLOUT, 2) + UBound(ItemLoopIns, 2)
                NewCol = UBound(FL, 1)
                ReDim Preserve FLOUT(1 To NewCol, 1 To NewLig)
                ReDim Preserve Formats(1 To NewLig)
                For cmpt2 = LBound(ItemLoopIns, 2) To UBound(ItemLoopIns, 2)
                    ColIter = 0
                    For cmpt1 = LBound(ItemLoopIns, 1) To UBound(ItemLoopIns, 1)
                        ColIter = ColIter + 1
                        FLOUT(ColIter, UBound(FLOUT, 2) - UBound(ItemLoopIns, 2) + cmpt2) = ItemLoopIns(cmpt1, cmpt2)
                    Next
                    LigIter = LigIter + 1
                Next
            Next
            ' On passe les lignes qui ont été étendues
            Nolig = Nolig + EndLoop - DebLoop
        Else
            If Nolig = 1 Then
                NewLig = 1
            Else
                NewLig = UBound(FLOUT, 2) + 1
            End If
            NewCol = UBound(FLOUT, 1)
            ReDim Preserve FLOUT(1 To NewCol, 1 To NewLig)
            For ColIter = LBound(FL, 1) To UBound(FL, 1)
                FLOUT(ColIter, NewLig) = FL(ColIter, Nolig)
            Next
        End If
    Next
    Deplier = FLOUT
End Function
Sub StructurationNomenclature(FLOUT As Worksheet, FLSKK As Worksheet)
    Dim Cell As Range, NoCol As Integer, Nolig As Long, NoSai As Long
    Dim derlig As Long, dercol As Integer, var As Variant, fct As Variant
    Dim WrdArray() As String
    Dim niv1 As String, niv2 As String, niv3 As String, leafi As String
    Dim OldPere As String
    Dim OldOk As Boolean
    '''''On Error GoTo ErrorHandler
    'Alimentation de la nomenclature
    a0 = InitTime()
    derlig = Split(FLNOM.UsedRange.Address, "$")(4)
    LigDebNom = 0
    For Nolig = 1 To derlig
        If FLNOM.Cells(Nolig, 1) = "*" Then
            LigDebNom = Nolig + 1
            Exit For
        End If
    Next
    derlig = Split(FLSKK.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLSKK.UsedRange.Address, "$")(3)).Column
    LigDebSSK = 0
    For Nolig = 1 To derlig
        If FLSKK.Cells(Nolig, 1) = "*" Then
            LigDebSSK = Nolig + 1
        End If
        If Left(FLSKK.Cells(Nolig, 1), 3) = "END" Then
            LigFinSSK = Nolig
        End If
    Next
    NbEqu = derlig - LigDebSSK
    Call EcritureInput(Etape, "" & NbEqu)
    Dim ENDSSK As Integer
    For NoCol = 1 To dercol
        var = FLSKK.Cells(1, NoCol)
        If Left(var, 3) = "END" Then ENDSSK = NoCol
    Next
    ENDSSKALPHA = DecAlph(ENDSSK)
    'Set FLSAI = Worksheets("TERTIAIRE")
    ' On ajoute la ligne pour récupérer le format plus tard
    Tab1ToProcess = Application.Transpose(FLSKK.Range("A" & LigDebSSK & ":" & DecAlph(ENDSSK) & LigFinSSK))

    For cmpt2 = LBound(Tab1ToProcess, 2) To UBound(Tab1ToProcess, 2)
        Tab1ToProcess(UBound(Tab1ToProcess, 1), cmpt2) = Tab1ToProcess(UBound(Tab1ToProcess, 1), cmpt2) & "°" & (LigDebSSK + cmpt2 - 1)
    Next
    a1 = SetTime()
    
    Tab2ToProcess = Deplier(Tab1ToProcess)
    Tab3ToProcess = Deplier(Tab2ToProcess)
    a2 = SetTime()
    
    TabFinal = Application.Transpose(Tab3ToProcess)
    a3 = SetTime()
    For cmpt1 = LBound(TabFinal, 1) To UBound(TabFinal, 1)
        ColIter = 0
        For cmpt2 = LBound(TabFinal, 2) To UBound(TabFinal, 2)
            ColIter = ColIter + 1
            FLOUT.Cells(cmpt1, ColIter).VALUE = TabFinal(cmpt1, cmpt2)
        Next
        LigIter = LigIter + 1
    Next
    a4 = SetTime()

    FLOUT.Cells(1, UBound(TabFinal, 2)).VALUE = "END" & "°" & (ENDSSK - 1)
    FLOUT.Cells(UBound(TabFinal, 1), 1).VALUE = "END" & "°1"
    ' Remplacement des [REF]
    derlig = Split(FLOUT.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLOUT.UsedRange.Address, "$")(3)).Column
    NbFin = derlig
    ToReplace = ""
    Dim Contenu As String
    a5 = SetTime()
    For Nolig = 1 To derlig
        For NoCol = 1 To dercol
            Contenu = FLOUT.Cells(Nolig, NoCol)
            If InStr(Contenu, "[") <> 0 Then
                chaine = Mid(Contenu, InStr(Contenu, "[") + 1, InStr(Contenu, "]") - 2)
                Gauche = Left(Contenu, InStr(Contenu, "["))
                PosDroite = Len(Contenu) - InStr(Contenu, "]") + 1
                If PosDroite > -1 Then
                    Droite = Right(Contenu, PosDroite)
                    If chaine = "REF" Then
                        FLOUT.Cells(Nolig, NoCol).VALUE = Gauche & ToReplace & Droite
                    Else
                        ToReplace = chaine
                    End If
                End If
            End If
        Next
    Next
    a6 = SetTime()
    ' Remplacement des libellé pour l'instant à partir de l'id plus tard à partir du nom dans NOMENCLATURE
    For Nolig = 1 To derlig
        For NoCol = 1 To dercol
            Contenu = FLOUT.Cells(Nolig, NoCol)
            If InStr(Contenu, "[") <> 0 Then
                If InStr(Contenu, "].") = 0 Then
                    chaine = Mid(Contenu, InStr(Contenu, "[") + 1, InStr(Contenu, "]") - 2)
                    LibArray = Split(chaine, ">")
                    libelle = LibArray(UBound(LibArray))
                    Contenu = Replace(Contenu, chaine, libelle)
                    Contenu = Replace(Contenu, "[" & libelle & "]", libelle)
                    FLOUT.Cells(Nolig, NoCol).VALUE = Contenu
                End If
            End If
        Next
    Next
    a7 = SetTime()
    Dim NumFormat As Integer
    ' Détermination des quantités et tag sur la dernière colonne
    For Nolig = 1 To derlig
        SplitNum = Split(FLOUT.Cells(Nolig, dercol), "°")
        For NoCol = 1 To dercol - 1
            Contenu = FLOUT.Cells(Nolig, NoCol)
            If InStr(Contenu, "(t)") <> 0 Then
                If UBound(SplitNum) > 0 Then
                    NumFormat = SplitNum(1)
                    FLOUT.Cells(Nolig, dercol).VALUE = Contenu & "°" & NumFormat
                Else
                    FLOUT.Cells(Nolig, dercol).VALUE = Contenu
                End If
            End If
        Next
    Next
    
    Call EcritureResultats(Etape, CONTEXTE, "" & NbFin)
    'MsgBox GetTime()
    Exit Sub
errorHandler: Call ErrorToDo(Etape, CONTEXTE, "" & NbFin, Err)
End Sub
Function InitTime()
    ReDim TIMEARRAY(1 To 1)
    TIMEARRAY(1) = Timer
    InitTime = 0
End Function
Function SetTime() As String
    ReDim Preserve TIMEARRAY(1 To UBound(TIMEARRAY) + 1)
    TIMEARRAY(UBound(TIMEARRAY)) = Timer
    SetTime = 0
End Function
Function GetTime() As String
    ReDim Preserve TIMEARRAY(1 To UBound(TIMEARRAY) + 1)
    TIMEARRAY(UBound(TIMEARRAY)) = Timer
    Total = TIMEARRAY(UBound(TIMEARRAY)) - TIMEARRAY(LBound(TIMEARRAY))
    Dim res As String
    res = ""
    For No = 1 To UBound(TIMEARRAY) - 1
        res = res & "   e" & No & "=" & Int(100 * (TIMEARRAY(No + 1) - TIMEARRAY(No)) / Total) & "%"
    Next
    res = Int(1000 * Total) & "ms = " & res
    GetTime = res
End Function
Sub MiseAuFormatSaisie()
    Etape = 11
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "TERTIAIRE"
    Call MiseAuFormat(FLSAI, FLSSK)
End Sub
Sub MiseAuFormatHM()
    Etape = 7
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "Hypothèses Macro"
    Call MiseAuFormat(FLHMA, FL0HM)
End Sub
Sub MiseAuFormatCalcul()
    Etape = 16
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "calculs TERTIAIRE"
    Call MiseAuFormat(FLCAL, FLCSK)
End Sub
Function GetCoordMatrice(FLSOU As Worksheet) As Variant()
    'Renvoie un vecteur des coordonnées lignes et colonnes de FLSOU
    Dim derlig As Integer, dercol As Integer
    dercol = Columns(Split(FLTIM.UsedRange.Address, "$")(3)).Column
    CodeHis = "$h"
    CodeHisLast = "$h.last"
    CodeJal = "$j"
    CodePro = "$p"
    NbHis = 0
    NbHisLast = 1
    NbJal = 0
    NbPro = 0
    derlig = Split(FLTIM.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLTIM.Cells(Nolig, 1) = "*" Then
            LigDebTime = Nolig + 1
            Exit For
        End If
    Next
    For NoCol = 1 To dercol
        var = FLTIM.Cells(LigDebTime, NoCol)
        If var = "h" Then
            NbHis = NbHis + 1
        End If
        If var = "j" Then
            NbJal = NbJal + 1
        End If
        If var = "j" Or var = "c" Then
            NbPro = NbPro + 1
        End If
    Next
    Dim LigDebFin() As Integer
    LigDebFin = GetFirstLine(FLSOU)
    LigDebSOU = LigDebFin(0)
    LigFinSOU = LigDebFin(1)
    Dim ColDebFin() As Integer
    ColDebFin = GetFirstCol(FLSOU)
    ColDebSOU = ColDebFin(0)
    ColFinSOU = ColDebFin(1)
    derlig = LigFinSOU
    dercol = ColFinSOU
    Dim SOU() As Variant
    ReDim SOU(1 To derlig - LigDebSOU + 1, 1 To dercol)
    SOU = FLSOU.Range("A" & LigDebSOU & ":" & DecAlph(dercol) & derlig).VALUE
    Dim CIB() As Variant
    ReDim CIB(1 To derlig - LigDebSOU + 1, 1 To dercol)
    Dim ColCoord() As String
    ReDim ColCoord(1 To dercol - 1)
    Adresse = 0
    LastAdresse = 1
    For NoCol = 1 To UBound(ColCoord)
        enPLus = 1
        If SOU(UBound(SOU, 1), NoCol) = "$h" Then enPLus = NbHis
        If SOU(UBound(SOU, 1), NoCol) = "$j" Then enPLus = NbJal
        If SOU(UBound(SOU, 1), NoCol) = "$p" Then enPLus = NbPro
        If SOU(UBound(SOU, 1), NoCol) = "$h" Then enPLus = NbHis
        LastAdresse = Adresse + 1
        Adresse = Adresse + enPLus
        ColCoord(NoCol) = LastAdresse & ":" & Adresse
    Next
    Dim Cas As String
    Dim Cla As String
    Dim LigCoord() As String
    ReDim LigCoord(1 To derlig - LigDebSOU)
    Adresse = 0
    LastAdresse = 1
    Dim Hier As Integer
    Hier = 0
    For Nolig = 1 To UBound(LigCoord)
        enPLus = 1
        LastAdresse = Adresse + 1
        Cas = SOU(Nolig, UBound(SOU, 2))
        If InStr(Cas, "EACH") > 0 Then
            Cla = GetClasseGauche("EACH", Cas)
            If InStr(Cas, ")") > 0 Then
                If Hier = 0 Then
                    Hier = 0
                    enPLus = UBound(GetChildrenClasse(Cla))
                    Adresse = Adresse + enPLus
                    LigCoord(Nolig) = LastAdresse & ":" & Adresse
                Else
                    enPLus = UBound(GetChildrenClasse(Cla))
                    Adresse = Adresse + enPLus
                    LigCoord(Nolig) = LastAdresse & ":" & Adresse
                    If Hier > 1 Then
                        Adresse = Adresse + 1
                        For NoH = 2 To Hier
                            If NoH = 2 Then
                                AddLine = 1
                            Else
                                AddLine = 2
                            End If
                            LastAdresse = Adresse + AddLine
                            Adresse1 = LastAdresse - 1
                            LigCoord(Nolig - 1) = LigCoord(Nolig - 1) & "|" & Adresse1 & ":" & Adresse1
                            Adresse = LastAdresse + enPLus - 1
                            LigCoord(Nolig) = LigCoord(Nolig) & "|" & LastAdresse & ":" & Adresse
                        Next
                    End If
                    Hier = 0
                End If
            Else
                Hier = UBound(GetChildrenClasse(Cla))
                enPLus = 1
                Adresse = Adresse + enPLus
                LigCoord(Nolig) = LastAdresse & ":" & Adresse
            End If
        Else
            Adresse = Adresse + enPLus
            LigCoord(Nolig) = LastAdresse & ":" & Adresse
        End If
    Next
    Dim res() As Variant
    ReDim res(1 To 2)
    res(1) = LigCoord
    res(2) = ColCoord
    GetCoordMatrice = res
End Function
Sub MiseAuFormat(FLCIB As Worksheet, FLSOU As Worksheet)
' Copy de la mise en forme du squelette vers la feuille finale
    On Error GoTo errorHandler
    LigSou = GetFirstLine(FLSOU)
    LigDebSOU = LigSou(0)
    LastLigSOU = LigSou(1)
    ColSou = GetFirstCol(FLSOU)
    ColDebSOU = ColSou(0)
    LastColSOU = ColSou(1)
    DataPos = GetCoordMatrice(FLSOU)
    Call EcritureInput(Etape, "" & (LastLigSOU * LastColSOU))

    Dim InteriorColor() As Variant
    Dim FontBold() As Variant
    Dim FontFontStyle() As Variant
    Dim FontName() As Variant
    Dim FontItalic() As Variant
    Dim FontUnderline() As Variant
    Dim FontColor() As Variant
    Dim FontSize() As Variant
    Dim NumberFormat() As Variant
    ReDim InteriorColor(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontBold(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontFontStyle(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontName(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontItalic(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontUnderline(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontColor(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim FontSize(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    ReDim NumberFormat(1 To LastLigSOU - 1, 1 To LastColSOU - 1)
    Dim L1 As Integer
    Dim L2 As Integer
    Dim c1 As Integer
    Dim C2 As Integer
    
    For Nolig = LigDebSOU To LastLigSOU - 1
        For NoCol = 1 To LastColSOU - 1
            NbS = UBound(Split(DataPos(1)(Nolig - LigDebSOU + 1), "|"))
            For NbSI = 0 To NbS
                Coord = Split(DataPos(1)(Nolig - LigDebSOU + 1), "|")(NbSI)
                SplLig = Split(Coord, ":")
                splCol = Split(DataPos(2)(NoCol), ":")
                L1 = Val(SplLig(0))
                L2 = Val(SplLig(1))
                c1 = Val(splCol(0))
                C2 = Val(splCol(1))
                With FLSOU.Cells(Nolig, NoCol)
                    NbS = UBound(Split(SplLig(1), "|"))
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Interior.Color = .Interior.Color
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.Bold = .Font.Bold
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.FontStyle = .Font.FontStyle
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.NAME = .Font.NAME
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.Italic = .Font.Italic
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.Underline = .Font.Underline
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.Color = .Font.Color
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).Font.Size = .Font.Size
                    FLCIB.Range(DecAlph(c1) & L1 & ":" & DecAlph(C2) & L2).NumberFormat = .NumberFormat
                End With
            Next
        Next
    Next
    NbFin = LastLigSOU * LastColSOU
    Call EcritureResultats(Etape, CONTEXTE, "" & NbFin)
    Exit Sub
errorHandler: Call ErrorToDo(Etape, CONTEXTE, "" & NbFin, Err)
End Sub
Sub Check(Etape As Integer, Texte As String)
'   Ecriture du résultat intermédiaire de l'étape Etape
    For Nolig = LigActionControl To DerLigControl
        If FLCONTROL.Cells(Nolig, 1) = Etape Then
            ''If FLCONTROL.Cells(NoLig, ColResultatControl + 1).Value = "" Then FLCONTROL.Cells(NoLig, ColResultatControl + 1) = Texte
            ''Else: FLCONTROL.Cells(NoLig, ColResultatControl + 1) = FLCONTROL.Cells(NoLig, ColResultatControl + 1) & " " & Texte
        End If
    Next
End Sub
Sub EcritureResultats(Etape As Integer, res As String, Texte As String)
'   Ecriture du résultat de l'étape Etape
    'Application.Calculation = xlCalculationAutomatic
    'Application.ScreenUpdating = True
    For Nolig = LigActionControl To DerLigControl
        If FLCONTROL.Cells(Nolig, 1) = Etape Then
            FLCONTROL.Cells(Nolig, COLOK) = "ok"
            FLCONTROL.Cells(Nolig, COLRESULTATNB) = Texte
            FLCONTROL.Cells(Nolig, COLRESULTAT) = res
            FLCONTROL.Cells(Nolig, 2).Interior.ColorIndex = 4
            FLCONTROL.Cells(Nolig, COLTIME) = Timer - START
        End If
    Next
End Sub
Sub EcritureInput(Etape As Integer, Texte As String)
'   Ecriture du résultat de l'étape Etape
    For Nolig = LigActionControl To DerLigControl
        If FLCONTROL.Cells(Nolig, 1) = Etape Then
            FLCONTROL.Cells(Nolig, COLINPUTNB) = Texte
        End If
    Next
End Sub
Sub DebutEtape(Etape As Integer)
'   Initialisation de l'étape Etape
    START = Timer
    'Application.Calculation = xlCalculationManual
    'Application.ScreenUpdating = False
    For Nolig = LigActionControl To DerLigControl
        If FLCONTROL.Cells(Nolig, 1) = Etape Then
            FLCONTROL.Cells(Nolig, 2).Interior.Color = RGB(255, 140, 0)
        End If
    Next
End Sub
Sub StructurationTimeSaisie()
    Etape = 9
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "TERTIAIRE"
    Call StructurationTime(FLSAI)
End Sub
Sub StructurationTimeHM()
    Etape = 5
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "Hypothèses Macro"
    Call StructurationTime(FLHMA)
End Sub
Sub StructurationTimeCalcul()
    Etape = 13
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "calculs TERTIAIRE"
    Call StructurationTime(FLCAL)
End Sub
Sub StructurationTime(FLENC As Worksheet)
Dim DoneHis As Boolean, DoneJal As Boolean, DonePro As Boolean, MatchTime As Boolean
Dim HisTime() As Integer
Dim JalTime() As Integer
Dim ProTime() As Integer
Dim TimType() As String
'Dim FLTIM As Worksheet
Dim derlig As Long, dercol As Integer, var As Variant
Dim NbHis As Long, NbJal As Long, NbPro As Long, MaxCol As Long
Dim CodeHis As String, CodeJal As String, CodePro As String
Dim Cell As Range, NoCol As Integer, Nolig As Long, NoSai As Long

Dim WrdArray() As String
Dim Pere As String
Dim Sous As String
Dim OldPere As String
Dim OldOk As Boolean
Dim SousSous As String
    On Error GoTo errorHandler
    a0 = InitTime()
    dercol = Columns(Split(FLTIM.UsedRange.Address, "$")(3)).Column
    CodeHis = "$h"
    CodeHisLast = "$h.last"
    CodeJal = "$j"
    CodePro = "$p"
    NbHis = 0
    NbHisLast = 1
    NbJal = 0
    NbPro = 0
    derlig = Split(FLTIM.UsedRange.Address, "$")(4)
    For Nolig = 1 To derlig
        If FLTIM.Cells(Nolig, 1) = "*" Then
            LigDebTime = Nolig + 1
            Exit For
        End If
    Next
    For NoCol = 1 To dercol
        var = FLTIM.Cells(LigDebTime, NoCol)
        If var = "h" Then
            NbHis = NbHis + 1
        End If
        If var = "j" Then
            NbJal = NbJal + 1
        End If
        If var = "j" Or var = "c" Then
            NbPro = NbPro + 1
        End If
    Next
    ReDim HisTime(NbHis)
    ReDim JalTime(NbJal)
    ReDim ProTime(NbPro)
    ReDim TimType(NbPro)
    NbHis = 0
    NbJal = 0
    NbPro = 0
    For NoCol = 1 To dercol
        var = FLTIM.Cells(LigDebTime, NoCol)
        If var = "h" Then
            NbHis = NbHis + 1
            HisTime(NbHis - 1) = FLTIM.Cells(LigDebTime + 1, NoCol)
        End If
        If var = "j" Then
            NbJal = NbJal + 1
            JalTime(NbJal - 1) = FLTIM.Cells(LigDebTime + 1, NoCol)
        End If
        If var = "j" Or var = "c" Then
            NbPro = NbPro + 1
            ProTime(NbPro - 1) = FLTIM.Cells(LigDebTime + 1, NoCol)
            TimType(NbPro - 1) = var
        End If
    Next
    derlig = Split(FLENC.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLENC.UsedRange.Address, "$")(3)).Column
    Call EcritureInput(Etape, "" & dercol - 1)
    a1 = SetTime()
    DoneHis = True
    DoneJal = True
    DonePro = True
    MaxCol = 0
    'Extension des colonnes selon le type h, j, p, h.last...
    Dim TypTime() As String
    ReDim TypTime(dercol)
    Dim PosTime() As Integer
    ReDim PosTime(dercol)
    For Nolig = 1 To derlig
        If Left(FLENC.Cells(Nolig, 1), 3) <> "END" Then NoEndLig = Nolig
    Next
    For Nolig = 1 To derlig
        'Repérage des indicateurs de temps
        If Left(FLENC.Cells(Nolig, 1), 3) <> "END" Then
            For NoCol = 1 To dercol
                var = FLENC.Cells(Nolig, NoCol)
                If Left(var, 1) = "$" Then
                    TypTime(NoCol) = var
                    If var = CodeHis Then PosTime(NoCol - 1) = NbHis - 1
                    If var = CodeHisLast Then PosTime(NoCol - 1) = NbHisLast - 1
                    If var = CodeJal Then PosTime(NoCol - 1) = NbJal - 1
                    If var = CodePro Then PosTime(NoCol - 1) = NbPro - 1
                End If
            Next
        End If
    Next
    a2 = SetTime()
    Dim AddCol As Integer
    Dim PosColToAdd As Integer
    Dim NbColToInsert As Integer
    AddCol = 0
    PosColToAdd = 0
    For NoCol = 0 To dercol - 1
        NbColToInsert = PosTime(NoCol)
        PosColToAdd = NoCol + 1 + 1 + AddCol
        If NbColToInsert > 0 Then
            For NoColToInsert = 1 To NbColToInsert
                FLENC.Columns(PosColToAdd).Insert Shift:=xlToRight
                AddCol = AddCol + 1
                FLENC.Cells(NoEndLig + 1, PosColToAdd) = FLENC.Cells(NoEndLig + 1, PosColToAdd) & "°" & NoCol + 1
            Next
        Else
            FLENC.Cells(NoEndLig + 1, PosColToAdd) = FLENC.Cells(NoEndLig + 1, PosColToAdd) & "°" & NoCol + 2
        End If
    Next
    a3 = SetTime()
    'Alimentation des colonnes avec les dates
    derlig = Split(FLENC.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLENC.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To dercol
        var = FLENC.Cells(1, NoCol)
        If Left(var, 3) = "END" Then NbColOutput = NoCol
    Next
    For Nolig = 1 To derlig
        var = FLENC.Cells(Nolig, 1)
        If Left(var, 3) = "END" Then NbLigOutput = Nolig
    Next
    Dim CodCol() As String
    Dim ValCol() As String
    Dim TypCol() As String
    ReDim CodCol(1 To NbLigOutput)
    ReDim ValCol(1 To NbLigOutput)
    ReDim TypCol(1 To NbLigOutput)
    For Nolig = 1 To NbLigOutput - 1
        For NoCol = 1 To NbColOutput - 1
            var = FLENC.Cells(Nolig, NoCol)
            If Left(var, 1) = "$" Then
                If var = CodeHis Then
                    FLENC.Cells(Nolig, NoCol) = HisTime(0)
                    For NoColToInsert = 1 To NbHis - 1
                        FLENC.Cells(Nolig, NoCol + NoColToInsert) = HisTime(NoColToInsert)
                        ValCol(NoCol + NoColToInsert) = HisTime(NoColToInsert)
                        TypCol(NoCol + NoColToInsert) = "h"
                    Next
                    CodCol(NoCol) = CodeHis
                    ValCol(NoCol) = HisTime(0)
                    TypCol(NoCol) = "h"
                End If
                If var = CodeHisLast Then
                    FLENC.Cells(Nolig, NoCol) = HisTime(NbHis - 1)
                    CodCol(NoCol) = CodeHisLast
                    ValCol(NoCol) = HisTime(NbHis - 1)
                    TypCol(NoCol) = "h"
                End If
                If var = CodeJal Then
                    FLENC.Cells(Nolig, NoCol) = JalTime(0)
                    For NoColToInsert = 1 To NbJal - 1
                        FLENC.Cells(Nolig, NoCol + NoColToInsert) = JalTime(NoColToInsert)
                        ValCol(NoCol + NoColToInsert) = JalTime(NoColToInsert)
                        TypCol(NoCol + NoColToInsert) = "j"
                    Next
                    CodCol(NoCol) = CodeJal
                    ValCol(NoCol) = JalTime(0)
                    TypCol(NoCol) = "j"
                End If
                If var = CodePro Then
                    FLENC.Cells(Nolig, NoCol) = ProTime(0)
                    For NoColToInsert = 1 To NbPro - 1
                        FLENC.Cells(Nolig, NoCol + NoColToInsert) = ProTime(NoColToInsert)
                        ValCol(NoCol + NoColToInsert) = ProTime(NoColToInsert)
                        TypCol(NoCol + NoColToInsert) = TimType(NoColToInsert)
                    Next
                    CodCol(NoCol) = CodePro
                    ValCol(NoCol) = ProTime(0)
                    TypCol(NoCol) = TimType(0)
                End If
                FLENC.Cells(Nolig, NbColOutput).VALUE = "TIME" & FLENC.Cells(Nolig, NbColOutput).VALUE
            End If
        Next
    Next
    a4 = SetTime()
    For NoCol = 2 To NbColOutput - 1
        FLENC.Cells(NbLigOutput, NoCol) = TypCol(NoCol) & "$" & ValCol(NoCol) & FLENC.Cells(NbLigOutput, NoCol)
    Next
    NbFin = NbColOutput - 1
    
    Call EcritureResultats(Etape, CONTEXTE, "" & NbFin)
    'MsgBox GetTime()
    Exit Sub
errorHandler: Call ErrorToDo(Etape, CONTEXTE, "" & NbFin, Err)
End Sub
Sub StructurationNomenclatureSaisie()
    Etape = 8
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "TERTIAIRE"
    Call StructurationNomenclature(FLSAI, FLSSK)
End Sub
Sub StructurationNomenclatureHM()
    Etape = 4
    Call DebutEtape(Etape)
    Call SetFl
    Call DelHM
    CONTEXTE = "Hypothèses Macro"
    Call StructurationNomenclature(FLHMA, FL0HM)
End Sub
Sub StructurationNomenclatureCalcul()
    Etape = 12
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "calculs TERTIAIRE"
    Call StructurationNomenclature(FLCAL, FLCSK)
End Sub
Sub AlimentationOldDataSaisie()
    Etape = 10
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "TERTIAIRE"
    Call AlimentationOldData(FLOLD, FLSAI, Etape)
End Sub
Sub AlimentationOldDataHM()
    Etape = 6
    Call DebutEtape(Etape)
    Call SetFl
    CONTEXTE = "Hypothèses Macro"
    Call AlimentationOldData(FLOHM, FLHMA, Etape)
End Sub
Sub AlimentationOldData(VIEUX As Worksheet, NOUVEAU As Worksheet, Etape As Integer)
    Dim DerLigSAI As Long, DerColSAI As Integer, DerLigOLD As Long, DerColOLD As Integer
    ''''On Error GoTo ErrorHandler
    a1 = InitTime()
    '''DerLigOLD = Split(VIEUX.UsedRange.Address, "$")(4)
    '''DerColOLD = Columns(Split(VIEUX.UsedRange.Address, "$")(3)).Column
    '''DerLigSAI = Split(NOUVEAU.UsedRange.Address, "$")(4)
    '''DerColSAI = Columns(Split(NOUVEAU.UsedRange.Address, "$")(3)).Column
    '''For NoCol = 1 To DerColSAI
        '''Var = NOUVEAU.Cells(1, NoCol)
        '''If Left(Var, 3) = "END" Then ENDSAI = NoCol
    '''Next
    '''For NoCol = 1 To DerColOLD
        '''Var = VIEUX.Cells(1, NoCol)
        '''If Left(Var, 3) = "END" Then ENDCOL = NoCol
    '''Next
    Dim ENDCOL As Integer
    Dim ENDSAI As Integer
    ColVieux = GetFirstCol(VIEUX)
    ENDCOL = ColVieux(1)
    DerColOLD = ColVieux(1)
    LigVieux = GetFirstLine(VIEUX)
    DerLigOLD = LigVieux(1)
    ColNew = GetFirstCol(NOUVEAU)
    ENDSAI = ColNew(1)
    DerColSAI = ColNew(1)
    LigNew = GetFirstLine(NOUVEAU)
    DerLigSAI = LigNew(1)
    TabOld = VIEUX.Range("A1:" & DecAlph(ENDCOL) & DerLigOLD).VALUE
    TabNew = NOUVEAU.Range("A1:" & DecAlph(ENDSAI) & DerLigSAI).VALUE
    'MsgBox DerLigOLD & ":" & ENDCOL & "::" & DerLigSAI & ":" & ENDSAI
    Dim ContexteGrandPere As String
    Dim ContextePere As String
    Dim ContexteFils As String
    Dim ContexteGrandPereOld As String
    Dim ContextePereOld As String
    Dim ContexteFilsOld As String
    Dim VectDate() As String
    Dim VectDate2() As String
    NbFin = 0
'MsgBox NOUVEAU.Name
a1 = SetTime()
For NoLigSAI = 1 To DerLigSAI
    tag = 0
    '''If NOUVEAU.Cells(NoLigSAI, 1) = "" Then ContexteGrandPere = NOUVEAU.Cells(NoLigSAI + 1, 1).Value
    '''If NOUVEAU.Cells(NoLigSAI, 3).Value > 1999 And NOUVEAU.Cells(NoLigSAI, 3).Value < 2041 Then
    If TabNew(NoLigSAI, 1) = "" Then ContexteGrandPere = TabNew(NoLigSAI + 1, 1)
    If TabNew(NoLigSAI, 3) > 1999 And TabNew(NoLigSAI, 3) < 2041 Then
        '''ContextePere = NOUVEAU.Cells(NoLigSAI, 1)
        ContextePere = TabNew(NoLigSAI, 1)
        tag = 1
        For NoColOLD = 2 To ENDSAI - 1
            '''If NOUVEAU.Cells(NoLigSAI, NoColOLD).Value > 1999 And NOUVEAU.Cells(NoLigSAI, NoColOLD).Value < 2041 Then
            If TabNew(NoLigSAI, NoColOLD) > 1999 And TabNew(NoLigSAI, NoColOLD) < 2041 Then
                ReDim Preserve VectDate2(2 To NoColOLD)
'If CONTEXTE = "TERTIAIRE" And NoColOLD < 4 Then MsgBox NoColOLD & ":" & VIEUX.Cells(NoLigCol, NoColOLD).Value
                VectDate2(NoColOLD) = "" & TabNew(NoLigSAI, NoColOLD)
            End If
'If CONTEXTE = "TERTIAIRE" And NoColOLD = ENDSAI - 1 Then MsgBox NoLigCol & " tailles dates:" & UBound(VectDate) & ":" & Join(VectDate, " ")
        Next
    End If
    If Left(TabNew(NoLigSAI, ENDSAI), 1) = "[" And tag = 0 Then
        ContexteFils = TabNew(NoLigSAI, 1)
        For NoLigCol = 1 To DerLigOLD
            TagOld = 0
            If TabOld(NoLigCol, 1) = "" Then ContexteGrandPereOld = TabOld(NoLigCol + 1, 1)
            If (TabOld(NoLigCol, 3) > 1999 And TabOld(NoLigCol, 3) < 2041) Or (TabOld(NoLigCol, 2) > 1999 And TabOld(NoLigCol, 2) < 2041) Then
            'And VIEUX.Cells(NoLigCol, 3).Value < 2041 Then
                ContextePereOld = TabOld(NoLigCol, 1)
                TagOld = 1
                'ReDim VectDate(2 To ENDSAI - 1)
                For NoColOLD = 2 To ENDSAI - 1
                    If TabOld(NoLigCol, NoColOLD) > 1999 And TabOld(NoLigCol, NoColOLD) < 2041 Then
                        ReDim Preserve VectDate(2 To NoColOLD)
'If CONTEXTE = "TERTIAIRE" And NoColOLD < 4 Then MsgBox NoColOLD & ":" & VIEUX.Cells(NoLigCol, NoColOLD).Value
                        VectDate(NoColOLD) = "" & TabOld(NoLigCol, NoColOLD)
                        NbFin = NbFin + 1
                    End If
'If CONTEXTE = "TERTIAIRE" And NoColOLD = ENDSAI - 1 Then MsgBox NoLigCol & " tailles dates:" & UBound(VectDate) & ":" & Join(VectDate, " ")
                Next
'If CONTEXTE = "TERTIAIRE" Then MsgBox NoColOLD & ":" & UBound(VectDate)
'If UBound(VectDate) < 3 Then MsgBox "AlimentationOldData " & UBound(VectDate) & ":" & Join(VectDate, ".")
            End If
            If TabOld(NoLigCol, 1) = ContexteFils And ContextePereOld = ContextePere And ContexteGrandPereOld = ContexteGrandPere And TagOld = 0 Then
                For NoColOLD = 2 To UBound(VectDate) 'ENDSAI - 1
                    TabNew(NoLigSAI, NoColOLD) = TabOld(NoLigCol, NoColOLD)
                Next
                Exit For
            End If
        Next
    End If
'If NOUVEAU.Cells(NoLigSAI, 1) = "Ajout" And CONTEXTE = "TERTIAIRE" Then
    'MsgBox NoLigCol & " " & NOUVEAU.Cells(NoLigSAI, 1) & " " & NOUVEAU.Cells(NoLigSAI, 2) & ":" & NOUVEAU.Cells(NoLigSAI, 3) & ":" & Join(VectDate2, " ")
'End If
    If Left(TabNew(NoLigSAI, 2), 1) = "[" Then
        For NoCol = 2 To UBound(VectDate2) 'ENDSAI - 1
            TabNew(NoLigSAI, NoCol) = 1
        Next
    End If
    If Left(TabNew(NoLigSAI, 3), 1) = "[" Then
        For NoCol = 3 To UBound(VectDate2) 'ENDSAI - 1
            TabNew(NoLigSAI, NoCol) = 1
        Next
    End If
Next
    NOUVEAU.Range("A1:" & DecAlph(ENDSAI) & DerLigSAI).VALUE = TabNew
    Call EcritureInput(Etape, "" & NbFin)
    Call EcritureResultats(Etape, CONTEXTE, "" & NbFin)
    'MsgBox GetTime()
    Exit Sub
errorHandler: Call ErrorToDo(Etape, CONTEXTE, "" & NbFin, Err)
End Sub
Sub AlimentationValeurs()
    Etape = 14
    Call DebutEtape(Etape)
    a0 = InitTime()
    On Error GoTo errorHandler
    Call SetFl
    Dim PosFct As Integer
    Dim LigFct As Integer
    NOMBRE = 0
    NbEquations = 0
    derlig = Split(FLCAL.UsedRange.Address, "$")(4)
    dercol = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
    For NoCol = 1 To dercol
        var = FLCAL.Cells(1, NoCol)
        If Left(var, 3) = "END" Then PosFct = NoCol
    Next
    Dim fct As String
    a1 = SetTime()
    For Nolig = 1 To derlig - 1
        Var1 = FLCAL.Cells(Nolig, PosFct)
        If Left(FLCAL.Cells(Nolig, PosFct), 1) = "[" Then
            fct = Split(Var1, "°")(0)
            LigFct = Nolig
            NbEquations = NbEquations + 1
            Call ChercherData(fct, LigFct)
        End If
    Next
    Call EcritureResultats(Etape, "", "")
    FLCONTROL.Select
    Call EcritureInput(Etape, "" & NbEquations)
    Call EcritureResultats(Etape, "calculs TERTIAIRE", "" & NOMBRE)
    'MsgBox GetTime()
    Exit Sub
errorHandler: Call ErrorToDo(Etape, "calculs TERTIAIRE", "" & NOMBRE, Err)
End Sub
Sub ChercherData(fct As String, LigFct As Integer)
    Dim Chemin() As String
    Dim Position() As Integer
    Chemin() = Split(fct, ".")
    FctHis = GetFctHisInEqu(fct)
    FctJal = GetFctJalInEqu(fct) 'PROBLEM
    If FctHis <> "" Or FctJal <> "" Then
        DerLigCal = Split(FLCAL.UsedRange.Address, "$")(4)
        DerColCal = Columns(Split(FLCAL.UsedRange.Address, "$")(3)).Column
        For NoCol = 1 To DerColCal
            var = FLCAL.Cells(1, NoCol)
            If Left(var, 3) = "END" Then PosFctCal = NoCol
        Next
        For Nolig = 1 To DerLigCal
            var = FLCAL.Cells(Nolig, 1)
            If Left(var, 3) = "END" Then PosTimCal = Nolig
        Next
        Dim NumberCol As Integer
        Dim TypTimSpl() As String
        Dim TypTimSplCal() As String
        Dim FL As Worksheet
        Dim PosFctSai As Integer
        Dim Contenu As String
        Dim feuille As String
        If FctHis <> "" Then
            WS = Split(FctHis, "'")(0)
            fc = Split(FctHis, "'")(1)
            If WS = FLCAL.NAME Then
                feuille = ""
            Else
                feuille = "'" & WS & "'!"
            End If
            Set FL = Worksheets(WS)
            derlig = Split(FL.UsedRange.Address, "$")(4)
            dercol = Columns(Split(FL.UsedRange.Address, "$")(3)).Column
            For NoCol = 1 To dercol
                var = FL.Cells(1, NoCol)
                If Left(var, 3) = "END" Then PosFctSai = NoCol
            Next
            For Nolig = 1 To derlig
                var = FL.Cells(Nolig, 1)
                If Left(var, 3) = "END" Then PosTimSai = Nolig
            Next
            For Nolig = 1 To PosTimSai - 1
                var = FL.Cells(Nolig, PosFctSai).VALUE
                Contenu = Split(var, "°")(0)
                If Contenu = fc Then
                    For NoCol = 1 To PosFctSai - 1
                        TypTimSpl() = Split(FL.Cells(PosTimSai, NoCol), "$")
                        If UBound(TypTimSpl) > 0 Then
                        typ = TypTimSpl(0)
                        DAT = Split(TypTimSpl(1), "°")(0)
                        If typ = "h" Then
                            For NoColCal = 1 To PosFctCal - 1
                                TypTimSplCal() = Split(FLCAL.Cells(PosTimCal, NoColCal), "$")
                                If UBound(TypTimSplCal) > 0 Then
                                TypCal = TypTimSplCal(0)
                                DatCal = Split(TypTimSplCal(1), "°")(0)
                                If DAT = DatCal Then
                                    NumberCol = NoCol
                                    FLCAL.Cells(LigFct, NoColCal).formula = "=" & feuille & DecAlph(NumberCol) & Nolig
                                    NOMBRE = NOMBRE + 1
                                End If
                                End If
                            Next
                        End If
                        End If
                    Next
                End If
            Next
        End If
        If FctJal <> "" Then
            WS = Split(FctJal, "'")(0)
            fc = Split(FctJal, "'")(1)
            If WS = FLCAL.NAME Then
                feuille = ""
            Else
                feuille = "'" & WS & "'!"
            End If
            Set FL = Worksheets(WS)
            derlig = Split(FL.UsedRange.Address, "$")(4)
            dercol = Columns(Split(FL.UsedRange.Address, "$")(3)).Column
            For NoCol = 1 To dercol
                var = FL.Cells(1, NoCol)
                If Left(var, 3) = "END" Then PosFctSai = NoCol
            Next
            For Nolig = 1 To derlig
                var = FL.Cells(Nolig, 1)
                If Left(var, 3) = "END" Then PosTimSai = Nolig
            Next
            For Nolig = 1 To PosTimSai - 1
                var = FL.Cells(Nolig, PosFctSai).VALUE
                Contenu = Split(var, "°")(0)
                If Contenu = fc Then
                    For NoCol = 1 To PosFctSai - 1
                        TypTimSpl() = Split(FL.Cells(PosTimSai, NoCol), "$")
                        If UBound(TypTimSpl) > 0 Then
                        typ = TypTimSpl(0)
                        DAT = Split(TypTimSpl(1), "°")(0)
                        If typ = "j" Then
                            For NoColCal = 1 To PosFctCal - 1
                                TypTimSplCal() = Split(FLCAL.Cells(PosTimCal, NoColCal), "$")
                                If UBound(TypTimSplCal) > 0 Then
                                TypCal = TypTimSplCal(0)
                                DatCal = Split(TypTimSplCal(1), "°")(0)
                                If DAT = DatCal Then
                                    NumberCol = NoCol
                                    FLCAL.Cells(LigFct, NoColCal).formula = "=" & feuille & DecAlph(NumberCol) & Nolig
                                    NOMBRE = NOMBRE + 1
                                End If
                                End If
                            Next
                        End If
                        End If
                    Next
                End If
            Next
        End If
    End If
End Sub

