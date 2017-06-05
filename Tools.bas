Attribute VB_Name = "Tools"
Public Sub DisableExcel()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub
Public Sub DisableExcelSoft()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    ActiveSheet.DisplayPageBreaks = True
End Sub

Public Sub EnableExcel()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.Cursor = xlDefault
End Sub
Public Sub EnableExcelSoft()
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ActiveSheet.DisplayPageBreaks = True
    Application.Cursor = xlDefault
End Sub
Public Function addToList(list() As String, Item As String) As String()
    If UBound(list) = 0 Then
        ReDim list(1 To 1)
    Else
        ReDim Preserve list(1 To UBound(list) + 1)
    End If
    list(UBound(list)) = Item
    addToList = list
End Function
Public Function appendLists(list1() As String, list2() As String) As String()
    Dim lastUbound As Integer
    If UBound(list1) = 0 Then
        appendLists = list2
    End If
    If UBound(list2) = 0 Then
        appendLists = list1
    End If
    If UBound(list1) = 0 And UBound(list2) = 0 Then
        appendLists = list1
    End If
    lastUbound = UBound(list1)
    If UBound(list1) > 0 And UBound(list2) > 0 Then
        ReDim Preserve list1(1 To UBound(list1) + UBound(list2))
        For I = 1 + lastUbound To UBound(list1)
            list1(I) = list2(I - lastUbound)
        Next
        appendLists = list1
    End If
End Function
Public Function addToListBool(list() As Boolean, Item As Boolean) As Boolean()
    If UBound(list) = 0 Then
        ReDim list(1 To 1)
    Else
        ReDim Preserve list(1 To UBound(list) + 1)
    End If
    list(UBound(list)) = Item
    addToListBool = list
End Function
Public Sub onErrDo(msg As String, where As String)
    Call EnableExcel
    MsgBox msg & Chr(10) & g_WSerror & Chr(10) & where & ":" & Err.Number & vbLf & Err.Description
    g_WSerror = ""
End Sub

