' =============================================================================
' Description: Controls visibility of multiple FEMAP groups in the active view.
' Compatibility: Modern FEMAP only; the script shows an early compatibility message on N4W 2004.
' =============================================================================

Option Explicit

Const DBUG As Boolean = False
Const C_FE_OK As Long = -1
Const C_FE_NO_MEMORY As Long = 10
Const C_FT_GROUP As Long = 24
Const C_FCM_NORMAL As Long = 0
Private er As Long
Private groupStates() As Long
Private groupStatesOrg() As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long

    er = 0

    If App.Info_Version < 10# Then
        MsgBox "Groups_Visibility requires the modern FEMAP multi-group View API and is not portable to Nastran for Windows 2004."
        GoTo Cleanup
    End If

	Dim v As Object
	Set v = App.feView
	rc = v.Get(1)

	If Not SupportsMultiGroupView(v) Then
		MsgBox "Groups_Visibility requires the modern FEMAP multi-group View API and is not portable to Nastran for Windows 2004."
		GoTo Cleanup
	End If

    If Not DBUG Then On Error GoTo Cleanup

	Dim nGroups As Long
	Dim nGroupID As Variant
	Dim g As Object
	Set g = App.feGroup()
	Dim gset As Object
	Set gset = App.feSet()

	er = 1

	rc = v.GetMultiGroupList(nGroups, nGroupID)
	If (rc = C_FE_NO_MEMORY) Then
		er = -1
		GoTo Cleanup
	End If

	er = 2

	rc =gset.AddAll(C_FT_GROUP)
	If rc <> C_FE_OK Then
		er = -2
		GoTo Cleanup
	End If
	ReDim groupStates(gset.First() To gset.Last())
	Call SetToZero(groupStates)

	er = 3

	For i = 0 To nGroups - 1 Step 1
		rc = g.Get(Abs(nGroupID(i)))
		groupStates(g.ID) = nGroupID(i) / Abs(nGroupID(i))
		If rc <> C_FE_OK Then
			If DBUG Then Debug.Print "Group ID - " & g.ID() & vbTab & "Does Not Exist"
			er = -3
			'GoTo Cleanup
		Else
				If DBUG Then Debug.Print "Group ID - " & g.ID() & vbTab & "Group Name - " & g.title() & vbTab & "State: " & nGroupID(i)/Abs(nGroupID(i))
		End If
	Next

	ReDim groupStatesOrg(LBound(groupStates) To UBound(groupStates))
	Call CopyLongArray(groupStates, groupStatesOrg)

	er = 4

	If DBUG Then
		For i = LBound(groupStates) To UBound(groupStates)
			rc = g.Get(i)
			If rc = C_FE_OK Then
				Debug.Print "Group ID - " & i & vbTab & "Group Name - " & g.title() & vbTab & "State: " & groupStates(i)
			End If
		Next
	End If

	er = 5

	Begin Dialog UserDialog 430,161,"Groups Visibility Options",.GroupsVisibility ' %GRID:10,7,1,1
		GroupBox 10,0,280,161,"Toggle",.GroupBox1
		PushButton 20,21,120,21,"All Visible",.allvisible
		PushButton 20,49,120,21,"All Invisible",.allinvisible
		PushButton 20,77,120,21,"All Hidden",.allhidden
		PushButton 150,21,130,21,"Inverse Visibility",.inverse
		PushButton 20,105,120,21,"Toggle Up",.toggleup
		PushButton 20,133,120,21,"Toggle Down",.toggledown
		PushButton 150,49,130,21,"Toggle But Visible",.togglebone
		PushButton 150,77,130,21,"Toggle But Hidden",.togglebmone
		PushButton 150,105,130,21,"Leave Visible",.leavevisible
		PushButton 150,133,130,21,"Leave Hidden",.leavehidden
		GroupBox 300,0,130,105,"Set",.GroupBox2
		PushButton 310,21,110,21,"Select Visible",.selvis
		PushButton 310,49,110,21,"Select Invisible",.selinvis
		PushButton 310,77,110,21,"Select Hidden",.selhid
		OKButton 340,112,90,21
		CancelButton 340,140,90,21
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) <> -1 Then
		Call CopyLongArray(groupStatesOrg, groupStates)
	End If

	Call Preview(groupStates)

	er = 0
Cleanup:
	If Not DBUG Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set v = Nothing
	Set gset = Nothing
	Set g = Nothing
	Set App = Nothing
End Sub

Private Function SupportsMultiGroupView(ByRef v As Object) As Boolean
	Dim nGroups As Long
	Dim ids As Variant
	Dim rc As Long
	On Error Resume Next
	Err.Clear
	rc = v.GetMultiGroupList(nGroups, ids)
	If Err.Number = 0 Then
		SupportsMultiGroupView = True
	Else
		SupportsMultiGroupView = False
	End If
	Err.Clear
End Function

Private Sub CopyLongArray(ByRef src() As Long, ByRef dst() As Long)
	Dim i As Long
	For i = LBound(src) To UBound(src)
		dst(i) = src(i)
	Next
End Sub

Rem See DialogFunc help topic for more information.
Private Function GroupsVisibility(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem GroupsVisibility = True ' Prevent button press from closing the dialog box
		GroupsVisibility = True
		Select Case DlgItem$
		Case "OK"
			GroupsVisibility = False
		Case "Cancel"
			GroupsVisibility = False
		Case "allvisible"
			Call SetToOne(groupStates)
			Call Preview(groupStates)
		Case "allinvisible"
			Call SetToZero(groupStates)
			Call Preview(groupStates)
		Case "allhidden"
			Call SetToMinusOne(groupStates)
			Call Preview(groupStates)
		Case "inverse"
			Call InverseState(groupStates)
			Call Preview(groupStates)
		Case "toggleup"
			Call ToggleStateUp(groupStates)
			Call Preview(groupStates)
		Case "toggledown"
			Call ToggleStateDown(groupStates)
			Call Preview(groupStates)
		Case "togglebone"
			Call ToggleButOne(groupStates)
			Call Preview(groupStates)
		Case "togglebmone"
			Call ToggleButMinusOne(groupStates)
			Call Preview(groupStates)
		Case "leavevisible"
			Call LeaveJustOne(groupStates)
			Call Preview(groupStates)
		Case "leavehidden"
			Call LeaveJustMinusOne(groupStates)
			Call Preview(groupStates)
		Case "selvis"
			Call SelectVisibility(groupStates, 1)
			Call Preview(groupStates)
		Case "selinvis"
			Call SelectVisibility(groupStates, 0)
			Call Preview(groupStates)
		Case "selhid"
			Call SelectVisibility(groupStates, -1)
			Call Preview(groupStates)
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : GroupsVisibility = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Private Sub SetToZero(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		arr(i) = 0
	Next
End Sub

Private Sub SetToOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		arr(i) = 1
	Next
End Sub

Private Sub SetToMinusOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		arr(i) = -1
	Next
End Sub

Private Sub ToggleStateUp(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case 1
			arr(i) = -1
		Case Else
			arr(i) = arr(i) + 1
		End Select
	Next
End Sub

Private Sub ToggleStateDown(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case -1
			arr(i) = 1
		Case Else
			arr(i) = arr(i) - 1
		End Select
	Next
End Sub

Private Sub LeaveJustOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case -1
			arr(i) = 0
		Case 1
			'Do Nothing
		Case Else
			'Do Nothing
		End Select
	Next
End Sub

Private Sub LeaveJustMinusOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case -1
			'Do Nothing
		Case 1
			arr(i) = 0
		Case Else
			'Do Nothing
		End Select
	Next
End Sub

Private Sub ToggleButOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case -1
			arr(i) = 0
		Case 1
			'Do Nothing
		Case Else
			arr(i) = -1
		End Select
	Next
End Sub

Private Sub ToggleButMinusOne(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		Select Case arr(i)
		Case -1
			'Do Nothing
		Case 1
			arr(i) = 0
		Case Else
			arr(i) = 1
		End Select
	Next
End Sub

Private Sub InverseState(ByRef arr() As Long)
	Dim i As Long
	For i = LBound(arr) To UBound(arr) Step 1
		arr(i) = -1 * arr(i)
	Next
End Sub

Private Function CountValues(ByRef arr() As Long) As Long
	Dim i As Long
	Dim c As Long
	For i = LBound(arr) To UBound(arr) Step 1
		c = c + Abs(arr(i))
	Next
	CountValues = c
End Function

Private Function SelectVisibility(ByRef arr() As Long, ByVal state As Long) As Long
	Dim App As Object
	Set App = feGetObject()
	Dim rc As Long, i As Long, erval As Long

	Dim txt(-1 To 1) As String
	txt(-1) = "Select Groups to be Hidden:"
	txt(0) = "Select Groups to be Invisible:"
	txt(1) = "Select Groups to be Visible:"

	erval = -1

	If Not DBUG Then On Error GoTo Cleanup

	Dim gset As Object
	Set gset = App.feSet()

	Dim g As Object
	Set g = App.feGroup()

	'Portable single/multiple group selection uses Set.Select
	rc = gset.Select(C_FT_GROUP, True, txt(state))
	If (rc <> -1) Or (gset.Count() < 1) Then
		erval = -2
		GoTo Cleanup
	End If

	erval = -3

	rc = gset.Reset()
	For i = 0 To gset.Count() - 1 Step 1
		rc = g.Get(gset.Next())
		If rc = C_FE_OK Then
			arr(g.ID) = state
		Else
			erval = i
		End If
	Next

	If (erval > -1) Then
		GoTo Cleanup
	End If

	erval = -1
Cleanup:
On Error Resume Next
	Set g = Nothing
	Set gset = Nothing
	Set App = Nothing
	SelectVisibility = erval
End Function

Private Sub Redraw()
	Dim App As Object
	Set App = feGetObject()

	Call App.feViewRegenerate(1)
	Call App.feViewRedraw(1)

On Error Resume Next
	Set App = Nothing
End Sub

Private Sub Preview(ByRef arr() As Long)
	Dim App As Object
	Set App = feGetObject()
	Dim rc As Long
	Dim i As Long, j As Long

	Dim v As Object
	Set v = App.feView()

	rc = v.Get(1)

	Dim gsets() As Long
	Dim gnum As Long

	gnum = CountValues(arr)
	If gnum = 0 Then
		Call v.ClearMultiGroupList()
	Else
		ReDim gsets(gnum - 1)
		j = -1
		For i = LBound(arr) To UBound(arr) Step 1
			If (arr(i) <> 0) Then
				j = j+1
				gsets(j) = i * arr(i)
			End If
		Next

		Call v.SetMultiGroupList(True, gnum, gsets)
	End If
	rc = v.Put(1)
	Call v.Regenerate()
	Call v.Redraw()

On Error Resume Next
	Set v = Nothing
	Set App = Nothing
End Sub
