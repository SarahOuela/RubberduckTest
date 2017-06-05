Attribute VB_Name = "Language"
Private Declare PtrSafe Function APIMsgBox _
Lib "User32" Alias "MessageBoxW" _
(Optional ByVal hWnd As Long, _
Optional ByVal Prompt As Long, _
Optional ByVal Title As Long, _
Optional ByVal Buttons As Long) _
As Long
Sub msg(idMsg As String)
    '''On Error GoTo errorHandler
    Dim mesmsg() As Variant
    Dim derlig As Integer
    Dim col As Integer
    Dim colId As Integer
    Dim nt As Integer
    Dim ntt As Integer
    Dim tex As String
    Dim tex_t As String
    
    colId = 3
    'Détermine la dernière ligne renseignée de la feuille de calculs
    derlig = getDerLig(Worksheets(g_Language_Sheet))
    mesmsg = Worksheets(g_Language_Sheet).Range("A1:f" & derlig).VALUE
    If g_Language = "fr" Then col = 4
    If g_Language = "cn" Then col = 5
    If g_Language = "en" Then col = 6
    nt = 0
    ntt = 0
    For I = 1 To UBound(mesmsg, 1) Step 1
        If mesmsg(I, colId) = idMsg Then
            nt = I
        End If
        If mesmsg(I, colId) = idMsg & "_t" Then
            ntt = I
        End If
    Next I
    If nt > 0 Then
        tex = mesmsg(nt, col)
        If ntt > 0 Then
            tex_t = mesmsg(ntt, col)
        Else
            tex_t = "Info"
        End If
        APIMsgBox Prompt:=StrPtr(tex), Title:=StrPtr(tex_t), Buttons:=vbOKOnly
    End If
Exit Sub
errorHandler: Call onErrDo("Il y a des erreurs", "msg"): Exit Sub
End Sub

Sub AfficheMsg2(msg As String)
'1 langue
Dim mesmsg
Dim msgtxt As String
mesmsg = Worksheets("Msg_Textes").Range("A1").CurrentRegion
msgtxt = ""

For I = 1 To UBound(mesmsg, 1) Step 1
    MsgBox msg & Chr(10) & mesmsg(I, 1)
    If mesmsg(I, 1) = msg Then
        For j = 1 To UBound(mesmsg, 2) Step 1
            msgtxt = msgtxt & IIf(Len(msgtxt) > 0, ", ", "") & mesmsg(I, j)
        Next j
    End If
Next I
APIMsgBox Prompt:=StrPtr(msgtxt), Title:=StrPtr(mesmsg(2, 1)), Buttons:=vbOKOnly
End Sub
Sub testini1()

Call AfficheMsg(g_Language, "Bonjour")

End Sub

Sub testiniMulti()
'Toutes langues

    Call AfficheMsg2("Bonjour")

End Sub


