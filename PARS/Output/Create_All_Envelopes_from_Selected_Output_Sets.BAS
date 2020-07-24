Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
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

    rc = osets.SelectMultiID(FT_OUT_CASE, 1, "Select Output Sets to Process to Envelopes.")
	If rc = FE_CANCEL Or osets.Count <= 1 Then GoTo Cleanup

	For i = 0 To UBound(nm)
		Call App.feAppMessage(FCM_NORMAL, "Creating Envelope " & nm(i))
		rc = App.feOutputProcessEnvelopeFromSets( True, i, 2, False, 2, True, osets.ID, 0, 0 )
	Next

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set osets = Nothing
	Set App = Nothing
End Sub
