Option Explicit

Const DBUG As Boolean = False
Private elArray() As String
Private elset As Object
Private ndset(2) As Object
Private length(3) As Double
Private lenmult(2) As Double
Private grpset As Object
Private grpID As Long
Private tmp As Variant
Private focus As Boolean
Private Beams As Object
Private overwriteL As Boolean

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long
    Dim tmp As Variant
    Dim i As Long

    If DBUG Then
		On Error GoTo 0
	Else
		On Error GoTo ExitHere
	End If

	Set Beams = App.feSet()
	Dim grp As Object
	Set grp = App.feGroup()

	For i = 1 To 10000
		rc = Beams.Reload2(i)
		If rc = -1 Then
			Beams.Save = True
			rc = Beams.First()
			If grp.Get(rc) = -1 Then
				If Left$(grp.title, 4) = "BEAM" Then
					rc = -1
					Exit For
				End If
			End If
		End If
		rc = 0
	Next

	If rc <> -1 Then
		Set Beams = App.feSet()
	End If

	Beams.Save = True

    Dim v As Object
    Set v = App.feView

    Dim grs As Object
    Set grs = App.feSet()
    Dim grsview() As Variant
    Dim grsview2 As Variant

    rc = v.Get(1)
    rc =grs.Select(FT_GROUP, True, "Select groups to be viewed, select non = same view:")
    If rc <> -1 Then GoTo ExitHere

    If grs.Count() > 0 Then
	    ReDim grsview(grs.Count()-1)

		rc = grs.Reset()
		For i = 0 To grs.Count() - 1
			grsview(i) = grs.Next()
			If DBUG Then Debug.Print grsview(i)
		Next

	    rc = v.SetMultiGroupList(True, grs.Count(), grsview)
	    rc = v.Put(1)
	    rc = v.Visible(True)
	    rc = v.Regenerate()
	End If

    Set grpset = App.feSet()

	Set grp = App.feGroup()
	Dim txt As Object
	Set txt = App.feText()

	grpID = 1000
	rc = -1
	Do While rc =-1
		grpID = grpID + 1
		rc = grp.Get(grpID)
		If txt.Get(grpID) =-1 Then rc = -1
	Loop

	If DBUG Then
		On Error GoTo 0
	Else
		On Error GoTo ExitHere
	End If

	Set ndset(0) = App.feSet()
	Set ndset(1) = App.feSet()
	Set ndset(2) = App.feSet()
	Set elset = App.feSet()

	ReDim elArray(0) As String
	elArray(0) = "Push Add to Start.."
	lenmult(0) = 1
	lenmult(1) = 1
	lenmult(2) = 1

	Begin Dialog GroupElements 0,300,680,273,"Sort elements To beams",.GroupElementsBeams ' %GRID:10,7,1,1
		TextBox 90,11,90,21,.groupID
		TextBox 350,56,120,21,.ndsy
		TextBox 350,112,120,21,.ndslt
		TextBox 480,112,60,21,.lcrltm
		text 10,14,70,14,"Group ID:",.Text1
		GroupBox 10,42,240,224,"Elements list:",.elementslist
		ListBox 20,56,220,182,elArray(),.els
		text 20,241,140,14,"Cumulative Length:",.Text2
		TextBox 150,238,90,21,.Length
		text 350,35,90,14,"Nodes",.Text6
		text 480,35,90,14,"Multiplier",.Text7
		text 550,59,90,14,"length",.lcry
		text 550,88,90,14,"length",.lcrz
		text 550,116,90,14,"length",.lcrlt
		text 640,59,40,14,"mm",.Text9
		text 640,88,40,14,"mm",.Text10
		text 640,116,40,14,"mm",.Text11
		TextBox 480,84,60,21,.lcrzm
		TextBox 350,84,120,21,.ndsz
		TextBox 480,56,60,21,.lcrym
		PushButton 260,140,80,21,"Add",.Add
		PushButton 260,168,80,21,"Remove",.Remove
		PushButton 260,196,80,21,"Reset",.Reset
		PushButton 260,224,80,21,"Show",.Show
		OKButton 480,245,90,21
		CancelButton 580,245,90,21
		CheckBox 260,56,90,14,"y-y",.yy
		CheckBox 260,84,90,14,"z-z",.zz
		CheckBox 260,112,90,14,"LT",.lt
	End Dialog
	Dim dlg As GroupElements
	rc = -1
	dlg.groupID = CStr(grpID - 1)
	dlg.yy = True
	dlg.zz = False
	dlg.lt = False
	Do While rc <> 0
		focus = False
		overwriteL = False
		dlg.groupID = CStr(CLng(dlg.groupID) + 1)
		dlg.lcrym = CStr(lenmult(0))
		dlg.lcrzm = CStr(lenmult(1))
		dlg.lcrltm = CStr(lenmult(2))
		length(0) = 0
		length(1) = 0
		length(2) = 0
		length(3) = 0
		dlg.ndsy = ""
		dlg.ndsz = ""
		dlg.ndslt = ""
		rc = ndset(0).Clear()
		rc = ndset(1).Clear()
		rc = ndset(2).Clear()
		rc = elset.Clear()
		tmp = -1
		rc = Dialog(dlg)
		If rc = -1 Then
			'ReDim grsview2
			rc = v.GetMultiGroupList(i, grsview2)
			ReDim Preserve grsview2(i)
			grsview2(i) = -grpID
			rc = v.SetMultiGroupList(True, i + 1, grsview2)
		    rc = v.Put(1)
		    rc = v.Visible(True)
		    rc = v.Regenerate()
		End If
	Loop
	If (rc = 0)  Then GoTo ExitHere

 ExitHere:
 	On Error Resume Next
 	Set txt = Nothing
 	Set grp = Nothing
	Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function GroupElementsBeams(DlgItem$, Action%, SuppValue?) As Boolean
	Dim i As Long
	Select Case Action%
	Case 1 ' Dialog box initialization
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
		'Set tmp = SuppValue
		If DlgValue("yy") = 0 Then
			Call DlgEnable("ndsy", False)
			Call DlgEnable("lcrym", False)
			length(0) = 0
		Else
			Call DlgEnable("ndsy", True)
			Call DlgEnable("lcrym", True)
		End If

		If DlgValue("zz") = 0 Then
			Call DlgEnable("ndsz", False)
			Call DlgEnable("lcrzm", False)
			length(1) = 0
		Else
			Call DlgEnable("ndsz", True)
			Call DlgEnable("lcrzm", True)
		End If

		If DlgValue("LT") = 0 Then
			Call DlgEnable("ndslt", False)
			Call DlgEnable("lcrltm", False)
			length(2) = 0
		Else
			Call DlgEnable("ndslt", True)
			Call DlgEnable("lcrltm", True)
		End If
		Call DlgText("lcry", "0")
		Call DlgText("lcrz", "0")
		Call DlgText("lcrlt", "0")
		grpID = CLng(DlgText("groupID"))
	Case 2 ' Value changing or button pressed
		Rem GroupElementsBeams = True ' Prevent button press from closing the dialog box
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
		GroupElementsBeams = True
		Call feFemap.feViewRegenerate(0)
		Select Case DlgItem$
		Case "OK"
			Call CreateGroup()
			GroupElementsBeams = False
		Case "Cancel"
			GroupElementsBeams = False
		Case "yy"
			If SuppValue = 0 Then
				Call DlgEnable("ndsy", False)
				Call DlgEnable("lcrym", False)
				length(0) = 0
			Else
				Call DlgEnable("ndsy", True)
				Call DlgEnable("lcrym", True)
			End If
		Case "zz"
			If SuppValue = 0 Then
				Call DlgEnable("ndsz", False)
				Call DlgEnable("lcrzm", False)
				length(1) = 0
			Else
				Call DlgEnable("ndsz", True)
				Call DlgEnable("lcrzm", True)
			End If
		Case "lt"
			If SuppValue = 0 Then
				Call DlgEnable("ndslt", False)
				Call DlgEnable("lcrltm", False)
				length(2) = 0
			Else
				Call DlgEnable("ndslt", True)
				Call DlgEnable("lcrltm", True)
			End If
		Case "Add"
			Call AddElements()
			Call DlgListBoxArray("els", elArray)
			Call DlgText("length", CStr(Round(length(3),1)))
		Case "Reset"
			Call ResetElements()
			Call DlgListBoxArray("els", elArray)
			Call DlgText("length", CStr(Round(length(3),1)))
		Case "els"
			tmp = SuppValue?
			Call ShowElement(SuppValue?)
		Case "Remove"
			If tmp > -1 Then
				Call RemoveElement(tmp)
				Call DlgListBoxArray("els", elArray)
				Call DlgText("length", CStr(Round(length(3),1)))
			End If
		Case "Show"
			Call ShowElements()

		End Select
	Case 3 ' TextBox or ComboBox text changed
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
		Select Case DlgItem$
		Case "lcrym"
			If Not IsNumeric(DlgText("lcrym")) Then
				DlgText("lcrym", CStr(lenmult(0)))
				DlgText("lcry", CStr(Round(lenmult(0)*length(0),1)))
			Else
				lenmult(0) = CDbl(DlgText("lcrym"))
				DlgText("lcry", CStr(Round(lenmult(0)*length(0),1)))
			End If
		Case "lcrzm"
			If Not IsNumeric(DlgText("lcrzm")) Then
				DlgText("lcrzm", CStr(lenmult(1)))
				DlgText("lcrz", CStr(Round(lenmult(1)*length(1),1)))
			Else
				lenmult(1) = CDbl(DlgText("lcrzm"))
				DlgText("lcrz", CStr(Round(lenmult(1)*length(1),1)))
			End If
		Case "lcrltm"
			If Not IsNumeric(DlgText("lcrltm")) Then
				DlgText("lcrltm", CStr(lenmult(2)))
				DlgText("lcrlt", CStr(Round(lenmult(2)*length(2),1)))
			Else
				lenmult(2) = CDbl(DlgText("lcrltm"))
				DlgText("lcrlt", CStr(Round(lenmult(2)*length(2),1)))
			End If
		Case "groupID"
			If Not IsNumeric(DlgText("groupID")) Then
				Call DlgText("groupID", CStr(grpID))
			Else
				If feFemap.feGroup.Get(CLng(DlgText("groupID"))) = -1 Then
					Call DlgText("groupID", CStr(grpID))
				Else
					grpID = DlgText("groupID")
				End If
			End If
		Case "Length"
			If Not IsNumeric(DlgText("Length")) Then
				Call DlgText("Length", CStr(length(3)))
			Else
				length(3) = CDbl(DlgText("Length"))
				overwriteL = True
			End If
		End Select
	Case 4 ' Focus changed
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
		Select Case DlgItem$
		Case "ndsy"
			If SuppValue <> DlgNumber("ndsy") Then
				Call DlgText(DlgItem$, SelectNodes(0))
				length(0) = NodesDistance(0)
				If DlgText(DlgItem$) <> "" Then DlgText("lcry", CStr(Round(lenmult(0)*length(0),1)))
			End If
		Case "ndsz"
			If SuppValue <> DlgNumber("ndsz") Then
				Call DlgText(DlgItem$, SelectNodes(1))
				length(1) = NodesDistance(1)
				If DlgText(DlgItem$) <> "" Then DlgText("lcrz", CStr(Round(lenmult(1)*length(1),1)))
			End If
		Case "ndslt"
			If SuppValue <> DlgNumber("ndslt") Then
				Call DlgText(DlgItem$, SelectNodes(2))
				length(2) = NodesDistance(2)
				If DlgText(DlgItem$) <> "" Then DlgText("lcrlt", CStr(Round(lenmult(2)*length(2),1)))
			End If
		End Select

	Case 5 ' Idle
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
		Rem Wait .1 : GroupElementsBeams = True ' Continue getting idle actions
	Case 6 ' Function key
		If DBUG Then Debug.Print(DlgItem$ & vbTab & Action% & vbTab & SuppValue?)
	End Select
End Function

Private Sub AddElements()
	Dim rc As Long
	Dim i As Long
	Dim el As Object
	Set el = feFemap.feElem()
	Dim l As Double

	If elset.Count > 0 Then Call feFemap.feViewShow(8, elset.ID)

	rc = elset.Select(8, False, "Select elements in one BEAM:")
	If rc <> -1 Then Exit Sub

	ReDim Preserve elArray(elset.Count() - 1)
	rc = elset.Reset()
	For i = 1 To elset.Count()
		elArray(i-1) = elset.Next()
		rc = el.Get(elArray(i-1))
		rc = el.Length(l)
		If Not overwriteL Then length(3) = length(3) + l
	Next
	Set el = Nothing
End Sub

Private Sub ResetElements()
	Dim rc As Long

	rc = elset.Reset()

	ReDim elArray(0) As String
	elArray(0) = "Push Add to Start.."
	length(3) = 0
	overwriteL = False
End Sub

Private Sub RemoveElement(ByVal index As Long)
	Dim rc As Long
	Dim i As Long
	Dim el As Object
	Set el = feFemap.feElem()
	Dim l As Double

	rc = elset.Remove(elArray(index))

	ReDim Preserve elArray(elset.Count())
	rc = elset.Reset()
	For i = 1 To elset.Count()
		elArray(i-1) = elset.Next()
		rc = el.Get(elArray(i-1))
		rc = el.Length(l)
		If Not overwriteL Then length(3) = length(3) + l
	Next
	Set el = Nothing
End Sub

Private Sub ShowElements()
	If elset.Count > 0 Then Call feFemap.feViewShow(8, elset.ID)
End Sub

Private Sub ShowElement(ByVal index As Long)
	If Not elArray(index) = "Push Add to Start.." Then Call feFemap.feViewShow(8, -CLng(elArray(index)))
End Sub

Private Function SelectNodes(ByVal index As Long) As String
	Dim rc As Long
	Dim ndID1 As Long
	Dim ndID2 As Long

	rc = ndset(index).Select(7, True, "Select two nodes for distance:")
	If (rc <> -1) Or (ndset(index).Count < 2) Then
		SelectNodes = ""
		Exit Function
	End If

	ndID1 = ndset(index).First()
	ndID2 = ndset(index).Last()

	SelectNodes = CStr(ndID1) & ", " & CStr(ndID2)
End Function

Private Function NodesDistance(ByVal index As Long) As Double
	Dim vecBase As Variant
	Dim vecdist As Variant
	Dim magDist As Double
	Dim ndID1 As Long
	Dim ndID2 As Long
	Dim rc As Long

	ndID1 = ndset(index).First()
	ndID2 = ndset(index).Last()

	rc = feFemap.feMeasureDistanceBetweenNodes(ndID1, ndID2, 0, 0, 0, vecBase, vecdist, magDist)

	NodesDistance = magDist
End Function

Private Sub CreateGroup()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long
	Dim data As String
	Dim i As Long

	Dim txt As Object
	Set txt = App.feText()

	Dim grp As Object
	Set grp = App.feGroup()

	Dim el As Object
	Set el = App.feElem()

	Dim prp As Object
	Set prp = App.feProp()
	Dim properties(9) As Double 'Area|Izz|Iyy|Iyz|J|Avzy|Avzz|Jw|ey|ez

	Dim bc As Object
	Set bc = App.feBeamCalculator()

	rc = elset.Reset()
	rc = elset.Next()
	rc =el.Get(rc)

	rc = prp.Get(el.propID)

	bc.Element = el.ID
	rc = bc.GetElementProperties(properties(0), properties(1), properties(2), properties(3), properties(4), properties(5), properties(6), properties(7), properties(8), properties(9))

	data = Join(Array(length(0) * lenmult(0), length(1) * lenmult(1), length(2) * lenmult(2), length(3)), vbTab)
	data = data & Chr(13) & vbLf  & "[" & prp.ID & "] - " & prp.title
	data = data & Chr(13) & vbLf

	For i = 0 To 9 Step 1
		data = data & properties(i) & vbTab
	Next
	data = Left$(data, Len(data) - 1)

	If DBUG Then rc = App.feAppMessage(FCM_NORMAL, data)

	txt.text = data
	rc = txt.Put(grpID)

	grp.title= "BEAM" & grpID & "[" & prp.ID & "] - " & prp.title
	rc = grp.SetAdd(8, elset.ID)
	rc = grp.Add(FT_PROP, prp.ID)
	rc = grp.Add(FT_TEXT, grpID)
	rc = grp.Put(grpID)

	rc = Beams.Add(grp.ID)

	rc = elset.Clear()
	rc = ndset(0).Clear
	rc = ndset(1).Clear
	rc = ndset(2).Clear
	ReDim elArray(0)
	elArray(0) = "Push Add to Start.."
	overwriteL = False

	On Error Resume Next
	Set bc = Nothing
	Set prp = Nothing
	Set el = Nothing
	Set grp = Nothing
	Set txt = Nothing
	Set App = Nothing
End Sub
