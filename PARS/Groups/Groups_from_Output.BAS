Option Explicit On
'Creates a Group of Elements based on their Output Value (either Nodal or Elemental).
'Selects the elements according to the input criteria - lower than limit, higher than limit,
'between limit values or out of limit values.

Const DBUG As Boolean = False
Private er As Long
Private maxval As Double
Private minval As Double
Private lval As Double
Private uval As Double
Private lvalp As Double
Private uvalp As Double
Private state As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = -1

    If Not DBUG Then On Error GoTo Cleanup

    Dim s As Object
    Set s = App.feSet()

    Dim g As Object
    Set g = App.feGroup()

    Dim oset As Object
    Set oset = App.feOutputSet()
    Dim osetID As Long

    Dim ovec As Object
    Set ovec = App.feOutput()
    Dim ovecID As Long

    Dim percent As Double

    Dim v As Object
    Set v = App.feView()

	rc = v.Get(1)

	If oset.Get(v.OutputSet) <> -1 Then
		rc = s.SelectID(FT_OUT_CASE, "Select OutputSet for script:", osetID)
		If rc <> -1 Then
			er = 0
			GoTo Cleanup
		End If
	Else
		osetID = v.OutputSet
	End If
    rc = oset.Get(osetID)
    rc = s.Clear()
    Call App.feAppMessage(FCM_NORMAL, "Selected Output Set: " & osetID)

    'rc = s.SelectOutputVectorID("Select OutputVector For script:", osetID, 0, 4, 0, True, ovecID)
    'Set ovec = oset.Vector(0)
	If (ovec.GetFromSet(oset.ID, v.ContourData) <> -1) Or (ovec.ID = 0) Then
		rc = s.SelectOutputVectorID("Select OutputVector For script:", osetID, 0, 4, 0, True, ovecID)
		If rc <> -1 Then
			er = 0
			GoTo Cleanup
		End If
	Else
		ovecID = v.ContourData
	End If

	rc = ovec.Get(ovecID)
	Call App.feAppMessage(FCM_NORMAL, "Selected Output Vector: " & ovecID)

	minval = ovec.minval()
	maxval = ovec.maxval()
	lval = minval
	uval = maxval
	If minval = 0 Then
		lvalp = 0
		state = 1
	Else
		lvalp = -1
		state = 3
	End If
	uvalp = 1


	Begin Dialog UserDialog 440,168,"Output to Groups",.ogroups ' %GRID:10,7,1,1
		GroupBox 10,7,430,133,"Limit Output Values for Vector: " & ovec.title,.GroupBox1
		text 20,28,200,14,"Min Value = " & Format$(minval, "##0.0E+00"),.Text1
		text 230,28,200,14,"Max Value = " & Format$(maxval, "##0.0E+00"),.Text2
		OptionGroup .Group1
			OptionButton 20,77,110,14,"Lower Than",.lt
			OptionButton 20,98,110,14,"Between",.bt
			OptionButton 20,119,110,14,"Higher Than",.ht
			OptionButton 20,56,130,14,"Out of Bounds",.ob
		CheckBox 180,56,110,14,"In Percent",.pc
		TextBox 180,77,90,21,.LBnd
		TextBox 330,77,90,21,.UBnd
		OKButton 250,147,90,21
		CancelButton 350,147,90,21
	End Dialog
	Dim dlg As UserDialog
	dlg.Group1 = state
	dlg.pc = 1
	dlg.LBnd = Format$(lvalp, "0.00")
	dlg.UBnd = Format$(uvalp, "0.00")
	rc = Dialog(dlg)
	If rc <> -1 Then
		er = 1
		GoTo Cleanup
	End If

	Dim firstID As Long, lastID As Long
	Dim e As Object
	Set e = App.feElem()
	Dim n As Object
	Set n = App.feNode()
	Dim neset As Object
	Set neset = App.feSet()

	Select Case ovec.location
	Case 7
		rc = n.First()
		firstID = n.ID()
		rc = n.Last()
		lastID = n.ID()
		rc = neset.AddAll(FT_NODE)
	Case 8
		rc = e.First()
		firstID = e.ID()
		rc = e.Last()
		lastID = e.ID()
		rc = neset.AddAll(FT_ELEM)
	End Select

	rc = neset.Reset()
	rc =s.Clear()
	g.ID = g.NextEmptyID()
	Select Case state
	Case 0
		Select Case ovec.location()
		Case 7
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) <= uval) Then rc = s.AddRule(rc, FGD_ELEM_BYNODE)
				End If
			Next
		Case 8
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) <= uval) Then rc = s.Add(rc)
				End If
			Next
		End Select
	Case 1
		Select Case ovec.location()
		Case 7
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) >= lval) And (ovec.Value(rc) <= uval) Then rc = s.AddRule(rc, FGD_ELEM_BYNODE)
				End If
			Next
		Case 8
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) >= lval) And (ovec.Value(rc) <= uval) Then rc = s.Add(rc)
				End If
			Next
		End Select
	Case 2
		Select Case ovec.location()
		Case 7
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) >= lval) Then rc = s.AddRule(rc, FGD_ELEM_BYNODE)
				End If
			Next
		Case 8
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) >= lval) Then rc = s.Add(rc)
				End If
			Next
		End Select
	Case 3
		Select Case ovec.location()
		Case 7
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) <= lval) Or (ovec.Value(rc) >= uval) Then rc = s.AddRule(rc, FGD_ELEM_BYNODE)
				End If
			Next
		Case 8
			For i = 0 To neset.Count() - 1 Step 1
				rc = neset.Next()
				If rc <> -1 Then
					If (ovec.Value(rc) <= lval) Or (ovec.Value(rc) >= uval) Then rc = s.Add(rc)
				End If
			Next
		End Select
	End Select

	rc = g.SetAdd(FT_ELEM, s.ID)
	Dim txt As String
	Select Case state
	Case 0
		txt = " (-inf, " & Format$(uval, "##0.0E+00") & ")"
	Case 1
		txt = " (" & Format$(uval, "##0.0E+00") & ", " & Format$(uval, "##0.0E+00") & ")"
	Case 2
		txt = " (" & Format$(uval, "##0.0E+00") & ", inf)"
	Case 3
		txt = " (-inf, " & Format$(lval, "##0.0E+00") & ")&(" & Format$(uval, "##0.0E+00") & ", inf)"
	End Select
	g.title = ovec.title & txt
	rc = g.Put(g.ID)
	If rc <> -1 Then
		er = 2
		GoTo Cleanup
	End If

	rc = v.SetMultiGroupList(True, 1, Array(g.ID))
	rc = v.Put(v.ID)
	rc = v.Regenerate()

	er = 0
Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set neset = Nothing
	Set n = Nothing
	Set e = Nothing
	Set s = Nothing
	Set oset = Nothing
	Set ovec = Nothing
	Set v = Nothing
	Set g = Nothing
	Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function ogroups(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem ogroups = True ' Prevent button press from closing the dialog box
		ogroups = True
		Select Case DlgItem$
		Case "OK", "Cancel"
			ogroups = False
		Case "Group1"
			state = SuppValue?
			Select Case SuppValue?
			Case 0
				Call DlgEnable("LBnd", False)
				Call DlgEnable("UBnd", True)
			Case 1, 3
				Call DlgEnable("LBnd", True)
				Call DlgEnable("UBnd", True)
			Case 2
				Call DlgEnable("LBnd", True)
				Call DlgEnable("UBnd", False)
			End Select
		Case "pc"
			If SuppValue? = 1 Then
				If minval = 0 Then
					Call DlgText("LBnd",  Format$((lval - minval) / (maxval - minval), "##0.00"))
					Call DlgText("UBnd",  Format$((uval - minval) / (maxval - minval), "##0.00"))
				Else
					Call DlgText("LBnd",  Format$(-1 + (lval - minval) / (maxval - minval) * 2, "##0.00"))
					Call DlgText("UBnd",  Format$(-1 + (uval - minval) / (maxval - minval) * 2, "##0.00"))
				End If
			Else
				If minval = 0 Then
					Call DlgText("LBnd",  Format$((maxval - minval) * (lvalp), "##0.0E+00"))
					Call DlgText("UBnd",  Format$((maxval - minval) * (uvalp), "##0.0E+00"))
				Else
					Call DlgText("LBnd",  Format$(minval + (maxval - minval) * (lvalp + 1) / 2, "##0.0E+00"))
					Call DlgText("UBnd",  Format$(minval + (maxval - minval) * (uvalp + 1) / 2, "##0.0E+00"))
				End If
			End If
		End Select
	Case 3 ' TextBox or ComboBox text changed
		If IsNumeric(CDbl(DlgText(DlgItem))) Then
			Select Case DlgValue("pc")
			Case 1
				If minval >= 0 Then
					If CDbl(DlgText(DlgItem)) < 0 Then
						DlgText(DlgItem, Format$(lvalp, "0.00"))
					ElseIf CDbl(DlgText(DlgItem)) > 1 Then
						DlgText(DlgItem, Format$(uvalp, "0.00"))
					Else
						Select Case DlgItem
						Case "LBnd"
							lvalp = CDbl(DlgText(DlgItem))
							lval = (maxval - minval) * (lvalp)
							DlgText(DlgItem, Format$(lvalp, "0.00"))
						Case "UBnd"
							uvalp = CDbl(DlgText(DlgItem))
							uval = (maxval - minval) * (uvalp)
							DlgText(DlgItem, Format$(uvalp, "0.00"))
						End Select
					End If
				Else
					If CDbl(DlgText(DlgItem)) < -1 Then
						DlgText(DlgItem, Format$(lvalp, "0.00"))
					ElseIf CDbl(DlgText(DlgItem)) > 1 Then
						DlgText(DlgItem, Format$(uvalp, "0.00"))
					Else
						Select Case DlgItem
						Case "LBnd"
							lvalp = CDbl(DlgText(DlgItem))
							lval = minval + (maxval - minval) * (lvalp + 1) / 2
							DlgText(DlgItem, Format$(lvalp, "0.00"))
						Case "UBnd"
							uvalp = CDbl(DlgText(DlgItem))
							uval = minval + (maxval - minval) * (uvalp + 1) / 2
							DlgText(DlgItem, Format$(uvalp, "0.00"))
						End Select
					End If
				End If
			Case 0
				If minval >= 0 Then
					If CDbl(DlgText(DlgItem)) < minval Then
						DlgText(DlgItem, Format$(lval, "##0.0E+00"))
					ElseIf CDbl(DlgText(DlgItem)) > maxval Then
						DlgText(DlgItem,Format$(uval, "##0.0E+00"))
					Else
						Select Case DlgItem
						Case "LBnd"
							lval = CDbl(DlgText(DlgItem))
							lvalp = (lval - minval) / (maxval - minval)
							DlgText(DlgItem, Format$(lval, "##0.0E+00"))
						Case "UBnd"
							uval = CDbl(DlgText(DlgItem))
							uvalp = (uval - minval) / (maxval - minval)
							DlgText(DlgItem,Format$(uval, "##0.0E+00"))
						End Select
					End If
				Else
					If CDbl(DlgText(DlgItem)) < minval Then
						DlgText(DlgItem, Format$(lval, "##0.0E+00"))
					ElseIf CDbl(DlgText(DlgItem)) > maxval Then
						DlgText(DlgItem,Format$(uval, "##0.0E+00"))
					Else
						Select Case DlgItem
						Case "LBnd"
							lval = CDbl(DlgText(DlgItem))
							lvalp = -1 + (lval - minval) / (maxval - minval) * 2
							DlgText(DlgItem, Format$(lval, "##0.0E+00"))
						Case "UBnd"
							uval = CDbl(DlgText(DlgItem))
							uvalp = -1 + (uval - minval) / (maxval - minval) * 2
							DlgText(DlgItem,Format$(uval, "##0.0E+00"))
						End Select
					End If
				End If
			End Select
		End If
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : ogroups = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
