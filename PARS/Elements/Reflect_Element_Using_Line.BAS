' =============================================================================
' Description: Reflects selected elements about a plane defined from a picked line.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FT_ELEM As Long = 8
Const C_MSG_NORMAL As Long = 0

'Reflects elements using a line perpendicular to the reflection plane.
'The reflection plane origin is the line midpoint.

Private er As Long
Const DBUG As Boolean = True

Sub Main()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long

    On Error GoTo Cleanup
    er = 0

    Dim elset As Object
    Set elset = App.feSet()

    Dim ndset As Object
    Set ndset = App.feSet()

    Dim ndID As Long

    Dim nd1 As Object
    Set nd1 = App.feNode()

    Dim nd2 As Object
    Set nd2 = App.feNode()

    Dim midnd(2) As Double
    Dim norm(2) As Double

    rc = elset.Select(8, True, "Select elements to be reflected:")
    If (rc <> -1) Or (elset.Count < 1) Then
        er = 1
        GoTo Cleanup
    End If

    rc = ndset.SelectID(7, "Select first node to create plane.", ndID)
    If (rc <> -1) Or (ndset.Count() < 1) Then
        er = 2
        GoTo Cleanup
    End If
    rc = nd1.Get(ndID)
    If rc <> -1 Then
        er = 2
        GoTo Cleanup
    End If

    rc = ndset.SelectID(7, "Select second node to create plane.", ndID)
    If (rc <> -1) Or (ndset.Count() < 2) Then
        er = 3
        GoTo Cleanup
    End If
    rc = nd2.Get(ndID)
    If rc <> -1 Then
        er = 3
        GoTo Cleanup
    End If

    midnd(0) = (nd1.x() + nd2.x()) / 2#
    midnd(1) = (nd1.y() + nd2.y()) / 2#
    midnd(2) = (nd1.z() + nd2.z()) / 2#

    norm(0) = nd2.x() - nd1.x()
    norm(1) = nd2.y() - nd1.y()
    norm(2) = nd2.z() - nd1.z()

    If Sqr(norm(0) * norm(0) + norm(1) * norm(1) + norm(2) * norm(2)) = 0# Then
        er = 4
        GoTo Cleanup
    End If

    rc = App.feGenerateReflect(C_FT_ELEM, elset.ID, midnd, norm, 0#, False)
    If rc <> -1 Then
        er = 5
        GoTo Cleanup
    End If

Cleanup:
    On Error Resume Next
    Set nd2 = Nothing
    Set nd1 = Nothing
    Set ndset = Nothing
    Set elset = Nothing
    If DBUG Then App.feAppMessage(C_MSG_NORMAL, "The script exited with code " & er)
    Set App = Nothing
End Sub
