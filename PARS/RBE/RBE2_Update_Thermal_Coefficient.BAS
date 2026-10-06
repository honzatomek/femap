' =============================================================================
' Description: Updates the rigid thermal-expansion coefficient for selected RBE2 elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Updates the rigid-element thermal expansion coefficient on selected RBE2 elements.
' RigidThermalExpansion is part of the legacy rigid-element API.

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
    Dim s As String
    Dim thexp As Double

    Set elset = App.feSet()
    Set el = App.feElem()

    On Error GoTo ErrorHandler

    rc = elset.Select(C_FT_ELEM, True, "Select RBE2 elements:")
    If rc <> C_FE_OK Then GoTo ExitHere

    s = InputBox("Input new Thermal Expansion Coefficient", "RBE2 Thermal Update", "1.2e-05")
    If s = "" Then GoTo ExitHere
    thexp = Val(s)

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And (Not el.RigidInterpolate) Then
            el.RigidThermalExpansion = thexp
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
