Option Explicit

'Launches Select Mesh Command for Selected Groups
Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    On Error GoTo CleanUp

    Dim gset As Object
    Set gset = App.feSet()
    Dim g As Object
    Set g = App.feGroup()

    Dim eset As Object
    'Set eset = App.feSet()

    rc = gset.SelectMultiID(FT_GROUP, 1, "Select groups for Select Mesh Command:")
    If (rc <> -1) Or (gset.Count < 1) Then
    	er = 1
    	GoTo CleanUp
    End If

    rc = gset.Reset()

	For i = 0 To gset.Count() - 1
		rc = g.Get(gset.Next())
		If rc <> -1 Then
			er = 2
			GoTo CleanUp
		End If
		Set eset = g.List(FGR_ELEM)
		If eset.Count() > 0 Then
			rc = g.SelectMesh(eset.ID())
			If rc <> -1 Then
				er = 3
				GoTo CleanUp
			End If
			rc = g.Put(g.ID)
			If rc <> -1 Then
				er = 4
				GoTo CleanUp
			End If
		End If
	Next

CleanUp:
	If DBUG Then App.feAppMessage(FCM_NORMAL, "The script exited with code " & er)
	On Error Resume Next
		Set App = Nothing
End Sub
