Attribute VB_Name = "HelpButton"
Sub Bouton123_Cliquer()
    MsgBox "ici"
End Sub
Sub Bouton13_Cliquer()
    MsgBox "la"
End Sub
Sub butParametersHelp_Click()
    msg ("modelFile")
    '''MsgBox "Le fichier du mod�le contient la description de :" & Chr(10) _
    '''& "- la structure du syst�me �tudi� (nomenclature)," & Chr(10) _
    '''& "- les quantit�s et �quations associ�es," & Chr(10) _
    '''& "- la description des feuilles � g�n�rer." & Chr(10) _
    ''', vbOKOnly, "Aide sur le fichier du mod�le"
End Sub
Sub butSourceHelp_Click()
    MsgBox "Le fichier source contient des feuilles" & Chr(10) _
    & "qui vont alimenter certaines cellules des feuilles cible." & Chr(10) _
    & "Une feuille source doit avoir le m�me nom que la feuille cible � alimenter." & Chr(10) _
    , vbOKOnly, "Aide sur le fichier source"
End Sub
Sub butTargetHelp_Click()
    MsgBox "Le fichier cible contient les feuilles devant �tre g�n�r�es," & Chr(10) _
    & "ainsi que des feuilles de donn�es devant �ventuellement �tre prises en compte." & Chr(10) _
    & "Une feuille cible doit porter le m�me nom que la feuille qui la d�crit dans le mod�le." & Chr(10) _
    , vbOKOnly, "Aide sur le fichier cible"
End Sub
Sub butOnprogressHelp_Click()
    msg ("enConstruction")
End Sub
