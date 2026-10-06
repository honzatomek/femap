' =============================================================================
' Description: Minimal portable PARS script template for N4W 2004 and modern FEMAP.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const DBUG As Boolean = True
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Dim rc As Long
    Dim i As Long

    Set App = feGetObject()
    er = 0

    If Not DBUG Then On Error GoTo Cleanup


Cleanup:
    On Error Resume Next
    If Not App Is Nothing Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    Set App = Nothing
End Sub
