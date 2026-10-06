' =============================================================================
' Description: Selects/highlights mesh belonging to user-selected FEMAP groups.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

'Launches Select Mesh Command for Selected Groups
Const DBUG As Boolean = False
Const C_FE_OK As Long = -1
Const C_FT_GROUP As Long = 24
Const C_FGR_ELEM As Long = 8
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long

    er = 0

    On Error GoTo CleanUp

    Dim gset As Object
    Set gset = App.feSet()
    Dim g As Object
    Set g = App.feGroup()

    Dim eset As Object
    'Set eset = App.feSet()

    rc = gset.Select(C_FT_GROUP, True, "Select groups for Select Mesh Command:")
    If (rc <> C_FE_OK) Or (gset.Count < 1) Then
    	er = 1
    	GoTo CleanUp
    End If

    rc = gset.Reset()

	For i = 0 To gset.Count() - 1
		rc = g.Get(gset.Next())
		If rc <> C_FE_OK Then
			er = 2
			GoTo CleanUp
		End If
		Set eset = g.List(C_FGR_ELEM)
		If eset.Count() > 0 Then
			rc = g.SelectMesh(eset.ID())
			If rc <> C_FE_OK Then
				er = 3
				GoTo CleanUp
			End If
			rc = g.Put(g.ID)
			If rc <> C_FE_OK Then
				er = 4
				GoTo CleanUp
			End If
		End If
	Next

CleanUp:
	If DBUG Then App.feAppMessage(C_FCM_NORMAL, "The script exited with code " & er)
	On Error Resume Next
		Set App = Nothing
End Sub
