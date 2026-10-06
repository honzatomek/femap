' =============================================================================
' Description: Scales all six stiffness components of selected spring-to-ground properties by one factor.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FE_FAIL As Long = 0
Const C_FE_CANCEL As Long = 2
Const C_FT_PROP As Long = 11
Const C_MSG_NORMAL As Long = 0

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long, i As Long
    Dim pset As Object
    Dim p As Object
    Dim pID As Long
    Dim scale As Double

    er = 0
    scale = 1#
    If Not DBUG Then On Error GoTo Cleanup

    Set pset = App.feSet()
    Set p = App.feProp()

    rc = pset.Select(C_FT_PROP, True, "Select Spring properties to scale:")
    Do While rc = -1
        If pset.Count() < 1 Then GoTo Cleanup

        rc = App.feGetReal("Input Stiffness Scale Value", -1.0E-16, 1.0E+16, scale)
        If rc = C_FE_CANCEL Then GoTo Cleanup

        rc = pset.Reset()
        pID = pset.Next()
        Do While pID > 0
            rc = p.Get(pID)
            If rc <> C_FE_FAIL Then
                If (p.type = 41) And (p.cbush = 1) Then
                    For i = 36 To 41 Step 1
                        p.pval(i) = p.pval(i) * scale
                    Next
                    rc = p.Put(pID)
                End If
            End If
            pID = pset.Next()
        Loop

        rc = pset.Select(C_FT_PROP, True, "Select Spring properties to scale:")
    Loop

Cleanup:
    On Error Resume Next
    Call App.feAppMessage(C_MSG_NORMAL, "The script exited with code: " & er)
    Set p = Nothing
    Set pset = Nothing
    Set App = Nothing
End Sub
