Option Explicit

Private ovects As Object
Private allvect As Boolean
Private osets() As Object
Private fact() As Double
Private num() As Long
Private CombsArray() As String
Private c As Long
Private pick As Long
Private env(2) As String
Private st As Long
Private clean As Boolean
Private StartTime As Double
Private com As Long
Private txt As String
Private sti As Boolean

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long
    Dim i As Long
	Dim nm As String
	Dim cmbs As Object
	Set cmbs = App.feSet()
	Dim todel As Object
	Set todel = App.feSet()
	Dim LdGroups As Long
	Dim comb() As Variant
	Dim groups As Long
	Dim osID As Long
	Dim oset As Object
	Set oset = App.feOutputSet()
	Dim envset As Object
	Set envset = App.feSet()
	Dim numcom As Long
	Dim multi As Long
	Dim suppress As Long

	StartTime = Timer

	env(0) = "Max "
	env(1) = "Min "
	env(2) = "MaxAbs "
	st = 1
	clean = True
	sti = False

	On Error GoTo ExitHere

    ReDim CombsArray(0) As String
    'ReDim osets(0)
    ReDim fact(0)
    fact(0) = 0.0
    allvect = False
    LdGroups = 1
    txt = ""
    CombsArray(0) = "Click Add to Start"
	Begin Dialog UserDialog 550,210,"Combinations via Groups v1.1",.combinations ' %GRID:10,7,1,1
		CheckBox 280,154,130,28,"Just Envelopes",.JustE
		CheckBox 280,126,130,28,"Set Info",.SetI
		GroupBox 0,7,270,203,"Combinations:",.GroupBox1
		ListBox 10,28,250,175,CombsArray(),.combs
		PushButton 420,42,120,21,"Add Group",.Add1
		PushButton 420,70,120,21,"Remove Group",.Remove1
		PushButton 420,98,120,21,"Change Factor",.Change1
		OKButton 420,154,120,21
		text 280,189,90,14,"Start With:",.Text1
		CancelButton 420,182,120,21
		TextBox 350,185,50,21,.StartWith1
		PushButton 280,42,90,21,"Up",.Up1
		PushButton 280,70,90,21,"Down",.Down1
		PushButton 420,126,120,21,"Change Group",.Change2
		TextBox 360,14,180,21,.name
		text 280,17,70,14,"Name:",.Text2
	End Dialog
	Dim dlg As UserDialog
	dlg.JustE = 1
	dlg.SetI = 0
	dlg.StartWith1 = CStr(st)
	dlg.name = txt
	rc = Dialog(dlg)
	If rc = 0 Then GoTo ExitHere
	If CombsArray(0) = "Click Add to Start" Then GoTo ExitHere

	Set ovects = App.feSet()
	rc = ovects.SelectMultiID(FT_OUT_DIR, osets(0).First(), "Select Output Vectors to combine:")
    If ovects.Count() = 0 Then
    	allvect = True
    End If

    suppress = App.DialogAutoSkipMsg
	App.DialogAutoSkipMsg = 1

    nm = Join(CombsArray, " + ")
    multi = 1

    For i = 0 To UBound(fact) - (st - 1) Step 1
		numcom = numcom + NumCombinations(UBound(fact) - st + 2, i + 1)
	Next

    'Group cleanup to create just envelopes
	For i = 0 To UBound(fact) Step 1
		If osets(i).Count() > 1 Then
			multi = 2
			Call MakeEnvelopes(osets(i), todel, i)
			LdGroups = 2
		End If
	Next

	'numcom = multi * numcom

	'Make combinations
	numcom = 0
	For i = 0 To UBound(fact) - (st - 1) Step 1
		numcom = numcom + NumCombinations(UBound(fact) - st, i + 1)
	Next

	Call App.feAppStatusShow(True, numcom)

	com = 0

	For i = 0 To UBound(fact) - (st - 1) Step 1
		'numcom = numcom + NumCombinations(UBound(fact) - st, i + 1)
		pick = i + 1
		Call Combine(pick, cmbs, LdGroups, allvect)
	Next

	'Cleanup
	rc = envset.Clear()

	For i = 0 To 2 Step 1
		Call App.feAppLock()
		rc = App.feOutputProcessEnvelopeFromSets(True, i, 2, False, True, sti, cmbs.ID, ovects.ID, 0)
		rc = oset.Last()
		osID = oset.ID
		oset.title = txt & env(i) & nm
		oset.notes = env(i) & "Envelope from Output Sets via femap API" & Chr(13) & Chr(10)  & env(i) & nm
		rc = envset.Add(osID)
		rc = oset.Put(osID)
		rc =App.feAppMessage(FCM_NORMAL, MinutesElapsed(StartTime) & " - " & txt & env(i) & nm)
		Call App.feAppUnlock()
	Next

	rc = todel.AddSet(cmbs.ID)

	If clean Then
		rc = App.feDelete(28, todel.ID)'
		rc = App.feRenumberOpt2(FT_OUT_CASE, envset.ID, 1, 0, 0, False, False, False, 0)
	End If

	App.DialogAutoSkipMsg = suppress

	rc = App.feAppMessage(FCM_NORMAL, "Script Lasted: " & MinutesElapsed(StartTime))

ExitHere:
	On Error Resume Next
	Call App.feAppStatusShow(False, numcom)
	Set oset = Nothing
	Set cmbs = Nothing
	Set todel = Nothing
	Set envset = Nothing
	For i = 0 To UBound(osets)
		Set osets(i) = Nothing
	Next
	Set ovects = Nothing
	Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function combinations(DlgItem$, Action%, SuppValue?) As Boolean
	Dim i As Long
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem combinations = True ' Prevent button press from closing the dialog box
		combinations = True
		Select Case DlgItem$
		Case "Cancel"
			combinations = False
		Case "Add1"
			Do While AddCombinationGroup()
				Call DlgListBoxArray("combs", CombsArray)
			Loop
		Case "Remove1"
			Call RemoveCombinationGroup(c)
			Call DlgListBoxArray("combs", CombsArray)
		Case "Change1"
			Call ChangeFactor(c)
			Call DlgListBoxArray("combs", CombsArray)
		Case "combs"
			c = SuppValue
		Case "OK"
			combinations = False
		Case "JustE"
			clean = Not clean
		Case "SetI"
			sti = Not sti
		Case "Up1"
			Call MoveUp()
			Call DlgListBoxArray("combs", CombsArray)
		Case "Down1"
			Call MoveDown()
			Call DlgListBoxArray("combs", CombsArray)
		Case "Change2"
			Call ChangeGroup()
			Call DlgListBoxArray("combs", CombsArray)
		End Select

	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
		Case "StartWith1"
			If Not IsNumeric(DlgText("StartWith1")) Then
				Call DlgText("StartWith1", CStr(st))
			Else
				st = CLng(DlgText("StartWith1"))
				If st < 1 Then st = 1
				Call DlgText("StartWith1", CStr(st))
			End If
		Case "name"
			txt = DlgText(DlgItem)
		End Select
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : combinations = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Private Sub ChangeGroup()
	Dim App As femap.model
    Set App = feFemap()
    Dim oset As Object
    Set oset = App.feSet()
    Dim rc As Long
	If (fact(0) = 0.0) Or (c > UBound(fact)) Or (c < 0) Then GoTo ExitHere

	Set oset = App.feSet()
	rc = App.feSelectOutputSets("Select Output sets of the Output Group " & c + 1, oset)
	If rc <> -1 Then GoTo ExitHere

	rc = osets(c).Clear()
	rc = osets(c).AddSet(oset.ID)

	CombsArray(c) = CStr(fact(c)) & " x ("
	rc = osets(c).Reset()
	rc = osets(c).Next()
	Do While rc <> 0
		CombsArray(c) = CombsArray(c) & rc & ","
		rc = osets(c).Next()
	Loop
	num(c) = osets(c).Count()
	CombsArray(c) = Left$(CombsArray(c), Len(CombsArray(c)) - 1)  & ")"

ExitHere:
	Set oset = Nothing
	Set App = Nothing
End Sub


Private Sub MoveUp()
	If (c <= 0) Or (fact(0) = 0.0) Then Exit Sub
	Dim tmp As Variant

	Set tmp = osets(c - 1)
	Set osets(c - 1) = osets(c)
	Set osets(c) = tmp

	tmp = fact(c - 1)
	fact(c - 1) = fact(c)
	fact(c) = tmp

	tmp = CombsArray(c - 1)
	CombsArray(c - 1) = CombsArray(c)
	CombsArray(c) = tmp

	tmp = num(c-1)
	num(c - 1) = num(c)
	num(c) = tmp

	c = c - 1
End Sub

Private Sub MoveDown()
	If (c >= UBound(fact)) Or (fact(0) = 0.0) Then Exit Sub
	Dim tmp As Variant

	Set tmp = osets(c + 1)
	Set osets(c + 1) = osets(c)
	Set osets(c) = tmp

	tmp = fact(c + 1)
	fact(c + 1) = fact(c)
	fact(c) = tmp

	tmp = CombsArray(c + 1)
	CombsArray(c + 1) = CombsArray(c)
	CombsArray(c) = tmp

	tmp = num(c + 1)
	num(c + 1) = num(c)
	num(c) = tmp

	c = c + 1
End Sub


Private Function AddCombinationGroup() As Boolean
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long
	Dim ub As Long
	Dim tmp As String
	Dim i As Long

	On Error Resume Next
	If fact(0) = 0.0 Then
		ReDim osets(0)
		ReDim fact(0)
		ReDim CombsArray(0)
		ReDim num(0)
		ub = 0
		tmp = "1,00"
	Else
		ub = UBound(fact) + 1
		ReDim Preserve osets(ub)
		ReDim Preserve fact(ub)
		ReDim Preserve CombsArray(ub)
		ReDim Preserve num(ub)
		tmp = CStr(fact(ub - 1))
	End If
	Set osets(ub) = App.feSet()
	rc = App.feSelectOutputSets("Select Output sets of the Output Group " & ub + 1, osets(ub))
	If rc <> - 1 Then
		ReDim Preserve osets(ub - 1)
		ReDim Preserve fact(ub - 1)
		ReDim Preserve num(ub - 1)
		ReDim Preserve CombsArray(ub - 1)
		AddCombinationGroup = False
		Exit Function
	End If

	tmp = InputBox("Input load factor for Group " & ub + 1, "Load Combinations", tmp)
	Do While Not IsNumeric(tmp)
		tmp = InputBox("Input load factor for Group " & ub + 1, "Load Combinations", tmp)
		If InStr(1, tmp, ".") <> 0 Then tmp = Replace(tmp, ".", ",")
		If tmp = "" Then tmp = "1.00"
		If CDbl(tmp) = 0 Then tmp = "1.00"
	Loop
	fact(ub) = CDbl(tmp)

	CombsArray(ub) = CStr(fact(ub)) & " x ("
	rc = osets(ub).Reset()
	rc = osets(ub).Next()
	Do While rc <> 0
		CombsArray(ub) = CombsArray(ub) & rc & ","
		rc = osets(ub).Next()
	Loop
	num(ub) = osets(ub).Count()
	CombsArray(ub) = Left$(CombsArray(ub), Len(CombsArray(ub)) - 1)  & ")"
	c = ub
	AddCombinationGroup = True

	Set App = Nothing
End Function

Private Sub RemoveCombinationGroup(Optional ByVal index As Long = -1)
	Dim i As Long
	Dim rc As Long
	If fact(0) = 0.0 Then Exit Sub

	If index < UBound(fact) Then
		For i = index To UBound(fact) - 1 Step 1
			rc = osets(i).Clear()
			rc = osets(i).AddSet(osets(i+1).ID)
			fact(i) = fact(i+1)
			CombsArray(i) = CombsArray(i+1)
		Next
	Else
		i = index
	End If
	ReDim Preserve osets(i- 1)
	ReDim Preserve fact(i- 1)
	ReDim Preserve CombsArray(i- 1)
	c = i - 1
End Sub

Private Sub ChangeFactor(ByVal index As Long)
	Dim tmp As String
	Dim rc As Long
	tmp = CStr(fact(index))
	tmp = InputBox("Input load factor for Group " & index + 1, "Load Combinations", tmp)
	Do While Not IsNumeric(tmp)
		tmp = InputBox("Input load factor for Group " & index + 1, "Load Combinations", tmp)
		If InStr(1, tmp, ".") <> 0 Then tmp = Replace(tmp, ".", ",")
		If tmp = "" Then tmp = "1.00"
		If CDbl(tmp) = 0 Then tmp = "1.00"
	Loop
	fact(index) = CDbl(tmp)
	CombsArray(index) = CStr(fact(index)) & " x ("
	rc = osets(index).Reset()
	rc = osets(index).Next()
	Do While rc <> 0
		CombsArray(index) = CombsArray(index) & rc & ","
		rc = osets(index).Next()
	Loop
	CombsArray(index) = Left$(CombsArray(index), Len(CombsArray(index)) - 1)  & ")"
End Sub


Private Sub MakeEnvelopes(ByRef s As Object, ByRef cmbs As Object, ByVal index As Long)
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long
    Dim eset As Object
    Set eset = App.feSet()
    Dim oset As Object
    Set oset = App.feOutputSet()
    Dim osID As Long
    Dim i As Long
	Call App.feAppLock()
	rc = s.Reset()
	ReDim sets(s.Count() - 1)

	For i = 0 To 1 Step 1
		rc = App.feOutputProcessEnvelopeFromSets(True, i, 2, False, True, False, s.ID, 0, 0)
		rc = oset.Last()
		osID = oset.ID
		oset.title = env(i) & CombsArray(index)
		oset.notes = env(i) & "Envelope from Output Sets via Femap API" & Chr(13) & Chr(10) & CombsArray(index)
		rc = eset.Add(osID)
		rc = oset.Put(osID)

		rc =App.feAppMessage(FCM_NORMAL, MinutesElapsed(StartTime) & " - " & env(i) & CombsArray(index))
	Next

	rc = s.Clear()
	rc = s.AddSet(eset.ID)
	rc = cmbs.AddSet(eset.ID)
	Call App.feAppUnlock()
ExitHere:
	On Error Resume Next
	Set oset = Nothing
	Set eset = Nothing
	Set App = Nothing
End Sub

Private Sub Combine(ByVal index As Long, ByRef cmbs As Object, Optional ByVal level As Long = 1, Optional ByVal av As Boolean = True)
	Dim App As femap.model
	Set App = feFemap()
	Dim oset As Object
	Set oset = App.feOutputSet()
	Dim rc As Long
	Dim comb() As Variant
	Dim hlp As Variant
	Dim osID As Long
	Dim i As Long, j As Long, k As Long
	Dim sets1() As Long
	Dim sets2() As Long
	Dim tmp As String
	Dim OutSet() As Long
	Dim OutVect() As Long
	Dim OutFact() As Double
	Dim ldgrp As Boolean

	ldgrp = False

	ReDim sets1(UBound(fact))
	ReDim sets2(UBound(fact))

	For i = 0 To UBound(fact) Step 1
		rc = osets(i).Reset()
		osID = osets(i).Next()
		If (osets(i).Count > 1) Then
			sets1(i) = osID
			sets2(i) = osets(i).Next()
		Else
			sets1(i) = osID
			sets2(i) = osID
		End If
	Next

	Call print_nCr(UBound(fact) + 1, index, comb, st)
	ReDim Preserve comb(c - 2)

	For i = 0 To UBound(comb) Step 1
		hlp = comb(i)
		Set oset = App.feOutputSet()
		ReDim OutSet((st - 1) + UBound(hlp))
		ReDim OutVect((st - 1)+ UBound(hlp))
		ReDim OutFact((st - 1) + UBound(hlp))

		For j = 0 To (st - 2) Step 1
			OutSet(j) = sets1(j)
			OutFact(j) = fact(j)
		Next
		For j = 0 To UBound(hlp) Step 1
			OutSet(j + st - 1) = sets1(hlp(j) - 1)
			OutFact(j + st - 1) = fact(hlp(j) - 1)
			If num(hlp(j) - 1) > 1 Then ldgrp = True
		Next
		tmp = "Comb "
		For j = 0 To UBound(OutSet) Step 1
			tmp = tmp & CStr(OutFact(j)) & "*" & CStr(OutSet(j)) & "+"
		Next
		tmp = Left(tmp, Len(tmp) - 1)
		Call App.feAppLock()
		Call CreateCombination(App, oset, cmbs, OutSet, OutFact, OutVect, av)
		rc = oset.Last()
		osID = oset.ID
		oset.title = tmp
		oset.notes = tmp
		rc = oset.Put(osID)
		rc =App.feAppMessage(FCM_NORMAL, MinutesElapsed(StartTime) & " - " & tmp)
		rc = cmbs.Add(osID)
		If ldgrp Then
			For j = 0 To (st - 2) Step 1
				OutSet(j) = sets2(j)
			Next
			For j = 0 To UBound(hlp) Step 1
				OutSet(j + st - 1) = sets2(hlp(j) - 1)
			Next
			tmp = "Comb "
			For j = 0 To UBound(OutSet) Step 1
				tmp = tmp & CStr(OutFact(j)) & "*" & CStr(OutSet(j)) & "+"
			Next
			tmp = Left(tmp, Len(tmp) - 1)
			Call CreateCombination(App, oset, cmbs, OutSet, OutFact, OutVect, av)
			rc = oset.Last()
			osID = oset.ID
			oset.title = tmp
			oset.notes = tmp
			rc = oset.Put(osID)
			rc =App.feAppMessage(FCM_NORMAL, MinutesElapsed(StartTime) & " - " & tmp)
			rc = cmbs.Add(osID)
		End If
		com = com + 1
		Call App.feAppStatusUpdate(com)
		Call App.feAppUnlock()
	Next

	Set oset = Nothing
	Set App = Nothing
End Sub

Private Sub CreateCombination(ByRef App As Variant, ByRef oset As Object, ByRef cmbs As Object, ByRef OutSet() As Long, ByRef OutFact() As Double, ByRef OutVect() As Long, Optional ByVal av As Boolean = True)
	Dim osID As Long
	Dim rc As Long
	Dim i As Long, j As Long

	If av Then
		osID = oset.NextEmptyID()
		'oset.ID = osID
		rc = App.feOutputProcessLinearCombination(av, 2, UBound(OutSet) + 1, OutFact, OutSet, 0, 0)
		'osID = oset.Last()
		'rc = cmbs.Add(osID)
	Else
		osID = oset.NextEmptyID()
		oset.ID = osID
		rc = oset.Put(osID)
		rc = ovects.Reset()
		For i = 0 To ovects.Count() - 1 Step 1
			OutVect(0) = ovects.Next()
			For j = 1 To UBound(OutSet) Step 1
				OutVect(j) = OutVect(0)
			Next
			rc = App.feOutputProcessLinearCombination(av, 2, UBound(OutSet) + 1, OutFact, OutSet, OutVect, osID)
			rc = cmbs.Add(osID)
		Next
	End If
End Sub

Function print_nCr(ByVal n As Integer, ByVal r As Integer, ByRef p() As Variant, Optional ByVal start As Long = 1)
  c = 1
  Dim tmp() As Long
  ReDim tmp(pick - 1)
  ReDim p(c - 1)
  p(c - 1) = tmp
  internal_print_nCr n, r, p, start, 1
End Function


Private Function internal_print_nCr(n As Integer, r As Integer, ByRef p() As Variant, Optional i As Integer, Optional l As Integer) As Integer

  ' n is the number of items we are choosing from
  ' r is the number of items to choose
  ' p is the upper corner of the output range
  ' i is the minimum item we are allowed to pick
  ' l is how many levels we are in to the choosing
  ' c is the complete set we are working on

  Dim tmp() As Long
  ReDim tmp(pick - 1)

  If n < 1 Or r > n Or r < 0 Then Err.Raise 1
  If i < 1 Then i = 1
  If l < 1 Then l = 1
  If c < 1 Then c = 1
  If r = 0 Then
    p(0)(0) = 1
    Exit Function
  End If

  Dim x As Integer
  Dim y As Integer

  For x = i To n - r + 1
    If r = 1 Then
      If c > 1 Then
        For y = 0 To l - 2
          If p(c - 1)(y) = 0 Then
            p(c - 1)(y) = p(c - 2)(y)
          End If
        Next
      End If
      p(c - 1)(l - 1) = x
      c = c + 1
      ReDim Preserve p(c - 1)
      p(c - 1) = tmp
    Else
      p(c - 1)(l - 1) = x
      Call internal_print_nCr(n, r - 1, p, x + 1, l + 1)
    End If
  Next

End Function

Private Function MinutesElapsed(ByVal StartTime As Double) As String
	MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
End Function

Private Function fac(ByVal n As Long) As Long
	Dim i As Long
	If n > 0 Then
		fac = n
		For i = n - 1 To 1 Step -1
			fac = fac * i
		Next
	ElseIf n = 0 Then
		fac = 1
	Else
		fac = 0
	End If
End Function

Private Function NumCombinations(ByVal n As Long, ByVal k As Long) As Long
	If (k <= n) And (k > 0) Then
		NumCombinations = fac(n) / (fac(k) * fac(n - k))
	Else
		NumCombinations = 0
	End If
End Function

