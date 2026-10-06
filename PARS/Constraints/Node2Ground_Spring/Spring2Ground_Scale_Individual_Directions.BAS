' =============================================================================
' Description: Scales the six spring-to-ground stiffness components by individually entered directional factors.
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
    Dim stiff_coeff(5) As Double

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Set pset = App.feSet()
    Set p = App.feProp()

    rc = pset.Select(C_FT_PROP, True, "Select Spring properties to scale:")
    Do While rc = -1
        If pset.Count() < 1 Then GoTo Cleanup

        rc = get_stiffness_coefficients(App, stiff_coeff)
        If rc = C_FE_CANCEL Then GoTo Cleanup

        rc = pset.Reset()
        pID = pset.Next()
        Do While pID > 0
            rc = p.Get(pID)
            If rc <> C_FE_FAIL Then
                If (p.type = 41) And (p.cbush = 1) Then
                    For i = 0 To 5 Step 1
                        p.pval(i + 36) = p.pval(i + 36) * stiff_coeff(i)
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

Function get_stiffness_coefficients(ByRef App As Object, ByRef stiffness_coefficients() As Double) As Long
    Dim rc As Long
    Dim vecBase1(2) As Double
    Dim vecBase2(2) As Double

    On Error GoTo Cleanup

    App.Info_DefaultX = 1#
    App.Info_DefaultY = 1#
    App.Info_DefaultZ = 1#
    rc = App.feCoordPick("Input Translational Stiffness Scale:", vecBase1)
    If rc = C_FE_CANCEL Then GoTo Cleanup

    App.Info_DefaultX = 1#
    App.Info_DefaultY = 1#
    App.Info_DefaultZ = 1#
    rc = App.feCoordPick("Input Rotational Stiffness Scale:", vecBase2)
    If rc = C_FE_CANCEL Then GoTo Cleanup

    stiffness_coefficients(0) = vecBase1(0)
    stiffness_coefficients(1) = vecBase1(1)
    stiffness_coefficients(2) = vecBase1(2)
    stiffness_coefficients(3) = vecBase2(0)
    stiffness_coefficients(4) = vecBase2(1)
    stiffness_coefficients(5) = vecBase2(2)

Cleanup:
    get_stiffness_coefficients = rc
End Function
