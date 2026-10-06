' =============================================================================
' Description: Sets all six DOFs active at the dependent/reference node of selected RBE3 elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Sets all six dependent/reference-node DOFs active for selected RBE3 elements.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim el As Object
    Dim elID As Long
    Dim rc As Long
    Dim i As Long

    Set elset = App.feSet()
    Set el = App.feElem()

    On Error GoTo ErrorHandler

    rc = elset.Select(C_FT_ELEM, True, "Select RBE3 elements:")
    If rc <> C_FE_OK Then GoTo ExitHere

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And el.RigidInterpolate Then
            For i = 0 To 5
                el.release(0, i) = 1
            Next
            rc = el.Put(el.ID)
        End If
        elID = elset.Next()
    Loop
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set el = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub
