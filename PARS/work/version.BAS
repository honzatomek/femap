Option Explicit


Sub Main
    Dim App As Object
    Dim rc as Long
    Dim errNumber as Long
    Dim errDescription As String

    On Error GoTo Cleanup

    Set App = feGetObject()

    MsgBox CStr(App.Info_Version)

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not App Is Nothing Then rc = App.feViewRegenerate(0)

    If errNumber <> 0 Then
        If Len(errDescription) > 100 Then errDescription = Left$(errDescription, 100)
        If Not App Is Nothing Then
            Call App.feAppMessage(3, "PARS Color error " & CStr(errNumber) & ": " & errDescription)
        End If
    End If

    Set App = Nothing
    On Error GoTo 0
End Sub

