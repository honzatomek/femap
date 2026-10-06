' =============================================================================
' Description: Creates result envelopes for selected output sets using modern FEMAP output-processing APIs.
' Compatibility: Modern FEMAP only; the script shows an early compatibility message on N4W 2004.
' =============================================================================

Option Explicit

Const C_FE_FAIL As Long = 0
Const C_FE_CANCEL As Long = 2
Const C_FT_OUT_CASE As Long = 28
Const C_FT_OUT_DIR As Long = 29
Const C_FCM_NORMAL As Long = 0

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    If Not RequireModernFemap(App, "Requires the post-2004 output-envelope API.") Then GoTo CompatibilityExit
    Dim rc As Long, i As Long

	Dim nm(0 To 2) As String
	nm(0) = "Max"
	nm(1) = "Min"
	nm(2) = "MaxAbs"

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim osets As Object
    Set osets = App.feSet()

    rc = osets.Clear()

    rc = osets.SelectMultiID(C_FT_OUT_CASE, 1, "Select Output Sets to Process to Envelopes.")
	If rc = C_FE_CANCEL Or osets.Count <= 1 Then GoTo Cleanup

	For i = 0 To UBound(nm)
		Call App.feAppMessage(C_FCM_NORMAL, "Creating Envelope " & nm(i))
		rc = App.feOutputProcessEnvelopeFromSets( True, i, 2, False, 2, True, osets.ID, 0, 0 )
	Next

Cleanup:
	Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set osets = Nothing
	Set App = Nothing

CompatibilityExit:
    On Error Resume Next
    Set App = Nothing
    Exit Sub

End Sub


Function RequireModernFemap(ByRef App As Object, ByVal featureText As String) As Boolean
    RequireModernFemap = False

    If App.Info_Version < 10# Then
        MsgBox "This script is not portable to Nastran for Windows 2004." & Chr(13) & Chr(10) & featureText
        Exit Function
    End If

    RequireModernFemap = True
End Function
