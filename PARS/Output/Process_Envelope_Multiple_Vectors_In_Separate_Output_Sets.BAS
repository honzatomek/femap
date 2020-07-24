Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim oset As Object
    Set oset = App.feOutputSet()

    Dim ovecs() As Variant
    Dim ovec As Object
    Set ovec = App.feOutput()

    Dim ss As Object
    Set ss = App.feSet()
	Dim sv As Object
	Set sv = App.feSet()

	Dim from_SetID() As Long
	Dim from_vectorID() As Long

	rc = ss.SelectMultiID(FT_OUT_CASE, 1, "Select Output Sets to process:")
	If (rc <> -1) Or (ss.Count() < 1) Then
		er = 1
		GoTo Cleanup
	End If

	er = -1
	rc = sv.SelectMultiID(FT_OUT_DIR, ss.First(), "Select Output Vectors for Enevelope:")
	If (rc <> -1) Or (sv.Count() < 1) Then
		er = 2
		GoTo Cleanup
	End If

	er = -2
	ReDim ovecs(sv.Count() - 1)
	ReDim from_SetID(sv.Count() - 1)
	ReDim from_vectorID(sv.Count() - 1)
	rc = sv.Reset()
	For i = 0 To sv.Count() - 1 Step 1
		Set ovecs(i) = App.feOutput
		rc = ovecs(i).Get(sv.Next())
		from_vectorID(i) = ovecs(i).ID()
	Next

	er = -3
	rc = ss.Reset()
	For i = 0 To ss.Count() - 1 Step 1
		rc = oset.Get(ss.Next())
		For j = 0 To sv.Count() - 1 Step 1
			from_SetID(j) = oset.ID()
		Next

		rc = App.feOutputProcessEnvelope(False, FOPE_MAX, FOPA_LOCATIONS, True, False, False, sv.Count(), from_SetID, from_vectorID, oset.ID())
		rc = App.feOutputProcessEnvelope(False, FOPE_MIN, FOPA_LOCATIONS, True, False, False, sv.Count(), from_SetID, from_vectorID, oset.ID())
		rc = App.feOutputProcessEnvelope(False, FOPE_MAXABS, FOPA_LOCATIONS, True, False, False, sv.Count(), from_SetID, from_vectorID, oset.ID())
	Next

er = 0
Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub
