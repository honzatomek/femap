Option Explicit

'Adds Mass Elements to selected nodes
Const DBUG As Boolean = False
Private er As Long
Private mass As Object

Sub Main()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long

	er = 0

	If DBUG Then
		On Error GoTo 0
	Else
		On Error GoTo Cleanup
	End If

	Dim ndset As Object
	Set ndset = App.feSet()

	Dim el As Object
	'Set el = App.feElem()

	Dim eset As Object
	Set eset = App.feSet()

	Set mass = App.feSet()

	Dim g As Object
	Set g = App.feGroup()

	rc = g.Get(App.Info_ActiveID(FT_GROUP))

	rc = mass.AddRule(FET_L_MASS , FGD_PROP_BYTYPE)
	If (rc <> -1) Or (mass.Count() < 1) Then
		er = 1
		GoTo Cleanup
	End If
	If DBUG Then Debug.Print "Number of mass elements: " & mass.Count()

	Dim pset As Object
	Set pset = App.feSet()
	Dim pID As Long
	Dim txt As String

	Do While ndset.Select(7, True, "Select nodes for Mass Element:") = -1
		i = -1
		Do While rc = -1
			i = i + 1
			rc = pset.Clear()
			rc = pset.SelectID(FT_PROP, GetMass(i), pID)
			If (rc <> -1) Or (pset.Count() < 1) Then
				er = 2
				GoTo Cleanup
			End If
			If pset.CountCommon(mass.ID) = 1 Then rc = 0
		Loop
		rc = g.Get(App.Info_ActiveID(FT_GROUP))
		rc = ndset.Reset()
		Set el = App.feElem()
		el.type = FET_L_MASS
		el.propID = pID
		el.color = 54
		el.topology = FTO_POINT1
		For i = 0 To ndset.Count() - 1 Step 1
			rc = ndset.Next()
			If rc = FE_FAIL Then
				er = 3
				GoTo Cleanup
			End If
			el.ID = el.NextEmptyID()
			el.Node(0) = rc
			rc = el.Put(el.ID)
			If rc <> -1 Then
				er = 4
				GoTo Cleanup
			End If
			rc = eset.Add(el.ID)
		Next
		'rc = App.feViewRegenerate(App.Info_ActiveID(FT_VIEW))
		If eset.Count() > 0 Then
			rc = g.SetAdd(FT_ELEM, eset.ID)
			rc = g.Put(g.ID)
			rc = eset.Clear()
		End If
		App.feViewRegenerate(0)
	Loop

Cleanup:
	If DBUG Then App.feAppMessage(FCM_NORMAL, "The script exited with code " & er)
	On Error Resume Next
		If (Not DBUG) And (er <> 0) Then
			If eset.Count > 0 Then	rc = App.feDelete(FT_ELEM, eset.ID)
		ElseIf DBUG And (er <> 0) Then
			If eset.Count > 0 Then
				rc = App.feViewShow(FT_ELEM, eset.ID)
			End If
		End If
		Set eset = Nothing
		Set el = Nothing
		Set pset = Nothing
		Set ndset = Nothing
		Set mass = Nothing
		Set App = Nothing
End Sub

Private Function GetMass(ByVal i As Long) As String
	Dim txt As String
	Dim j As Long

	Select Case i
	Case 0
		txt = "Select Mass Property:"
	Case 1
		txt = "Select MASS Property:"
	Case 2
		txt = "Only MASS Property can be selected:"
	Case 3 To 9
		txt = "MASS MASS MASS Property MUST be SELECTED!"
	Case 10
		txt = "Just enjoy yourself.. Best by selecting MASS Property:"
	Case 20,30,40,50
		txt = "This is already the " & i & "st try, press Cancel or select MASS Property:"
	Case Else
		txt = "["
		For j = 1 To mass.Count() Step 1
			txt = txt & mass.Next() & ","
		Next
		txt = Left(txt, Len(txt) - 1) & "]"
		txt = "Only following MASS Properties can be selected: " & txt
	End Select
	GetMass = txt
End Function
