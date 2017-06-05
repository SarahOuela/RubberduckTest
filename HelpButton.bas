Attribute VB_Name = "HelpButton"
Sub Bouton123_Cliquer()
    MsgBox "ici"
End Sub
Sub Bouton13_Cliquer()
    MsgBox "la"
End Sub
Sub butParametersHelp_Click()
    msg ("modelFile")
    '''MsgBox "Le fichier du modèle contient la description de :" & Chr(10) _
    '''& "- la structure du système étudié (nomenclature)," & Chr(10) _
    '''& "- les quantités et équations associées," & Chr(10) _
    '''& "- la description des feuilles à générer." & Chr(10) _
    ''', vbOKOnly, "Aide sur le fichier du modèle"
End Sub
Sub butSourceHelp_Click()
    MsgBox "Le fichier source contient des feuilles" & Chr(10) _
    & "qui vont alimenter certaines cellules des feuilles cible." & Chr(10) _
    & "Une feuille source doit avoir le même nom que la feuille cible à alimenter." & Chr(10) _
    , vbOKOnly, "Aide sur le fichier source"
End Sub
Sub butTargetHelp_Click()
    MsgBox "Le fichier cible contient les feuilles devant être générées," & Chr(10) _
    & "ainsi que des feuilles de données devant éventuellement être prises en compte." & Chr(10) _
    & "Une feuille cible doit porter le même nom que la feuille qui la décrit dans le modèle." & Chr(10) _
    , vbOKOnly, "Aide sur le fichier cible"
End Sub
Sub butOnprogressHelp_Click()
    msg ("enConstruction")
End Sub
