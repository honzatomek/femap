' =============================================================================
' Description: Sets spring-to-ground translational stiffnesses to 0.38, 0.62, and 1.00 times a user-entered reference value.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_FAIL As Long = 0
Const C_FE_CANCEL As Long = 2
Const C_FT_PROP As Long = 11

Sub Main
    Dim App As Object
    Dim pset As Object
    Dim p As Object
    Dim rc As Long
    Dim pID As Long
    Dim value As Double

    value = 1#
    Set App = feGetObject()
    Set pset = App.feSet()
    Set p = App.feProp()

    rc = pset.Select(C_FT_PROP, True, "Select Spring properties to set stiffness values:")
    Do While rc = -1
        If pset.Count() < 1 Then GoTo Cleanup

        rc = App.feGetReal("Input Reference Translational Stiffness", -1.0E-16, 1.0E+16, value)
        If rc = C_FE_CANCEL Then GoTo Cleanup

        rc = pset.Reset()
        pID = pset.Next()
        Do While pID > 0
            rc = p.Get(pID)
            If rc <> C_FE_FAIL Then
                If (p.type = 41) And (p.cbush = 1) Then
                    p.pval(36) = value * 0.38
                    p.pval(37) = value * 0.62
                    p.pval(38) = value
                    rc = p.Put(pID)
                End If
            End If
            pID = pset.Next()
        Loop

        rc = pset.Select(C_FT_PROP, True, "Select Spring properties to set stiffness values:")
    Loop

Cleanup:
    On Error Resume Next
    Set p = Nothing
    Set pset = Nothing
    Set App = Nothing
End Sub
