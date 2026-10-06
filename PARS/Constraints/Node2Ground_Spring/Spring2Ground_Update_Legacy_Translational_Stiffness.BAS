' =============================================================================
' Description: Applies the legacy spring-to-ground translational stiffness correction: swap K1/K2 and divide K1, K2, and K3 by six.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_FAIL As Long = 0
Const C_FT_PROP As Long = 11
Const C_MSG_NORMAL As Long = 0

Sub Main
    Dim App As Object
    Dim pset As Object
    Dim p As Object
    Dim rc As Long
    Dim pID As Long
    Dim oldK1 As Double

    Set App = feGetObject()
    Set pset = App.feSet()
    Set p = App.feProp()

    rc = pset.Select(C_FT_PROP, True, "Select Spring properties to update:")
    If (rc <> -1) Or (pset.Count() < 1) Then GoTo Cleanup

    rc = pset.Reset()
    pID = pset.Next()
    Do While pID > 0
        rc = p.Get(pID)
        If rc <> C_FE_FAIL Then
            If p.cbush = 1 Then
                oldK1 = p.pval(36)
                p.pval(36) = p.pval(37) / 6#
                p.pval(37) = oldK1 / 6#
                p.pval(38) = p.pval(38) / 6#
                rc = p.Put(pID)
            End If
        End If
        pID = pset.Next()
    Loop

Cleanup:
    On Error Resume Next
    Set p = Nothing
    Set pset = Nothing
    Set App = Nothing
End Sub
