Attribute VB_Name = "Module2"
Option Explicit

Const DBUG As Boolean = True
Private beams As Object

Sub Main()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long
	Dim i As Long, j As Long

	If DBUG Then
		On Error GoTo 0
	Else
		On Error GoTo ExitHere
	End If

	Set beams = App.feSet()
	Dim grp As Object
	Set grp = App.feGroup()

	For i = 1 To 10000
		rc = beams.Reload2(i)
		If rc = -1 Then
			beams.Save = True
			rc = beams.First()
			If grp.Get(rc) = -1 Then
				If Left$(grp.title, 4) = "BEAM" Then
					If DBUG Then Debug.Print grp.title
					rc = -1
					Exit For
				End If
			End If
		End If
		rc = 0
	Next

	If rc = 0 Then GoTo ExitHere

	Dim dt As Object
	Set dt = App.feDataTable()

	Set grp = App.feGroup()
	Dim txt As Object
	Set txt = App.feText()
	Dim tmp As Variant
	Dim tmp2 As Variant

	Dim s As Object
	Set s= App.feSet()

	Dim el As Object
	Set el = App.feElem()

	Dim p As Object
	Set p = App.feProp()

	Dim elID() As Long
	Dim propID() As Long
	Dim propname() As String
	Dim grpID() As Long
	Dim grpname() As String
	Dim length() As Double
	Dim lcry() As Double
	Dim lcrz() As Double
	Dim lcrlt() As Double
	Dim num As Long
	Dim num2 As Long

	num = -1
	num2 = -1

	rc = beams.Reset()
	For i = 0 To beams.Count() - 1 Step 1
		rc = grp.Get(beams.Next())
		rc = txt.Get(grp.ID)
		tmp = Split(txt.text, vbLf)
		tmp = Split(tmp(0), vbTab)

		num = num + grp.CountList(8)
		ReDim Preserve elID(num)
		ReDim Preserve propID(num)
		ReDim Preserve propname(num)
		ReDim Preserve grpID(num)
		ReDim Preserve grpname(num)
		ReDim Preserve length(num)
		ReDim Preserve lcry(num)
		ReDim Preserve lcrz(num)
		ReDim Preserve lcrlt(num)
		Set s = grp.List(8)
		rc =s.Reset()
		For j = 0 To grp.CountList(8) - 1 Step 1
			num2 = num2 + 1
			rc = el.Get(s.Next)
			rc = p.Get(el.propID)
			elID(num2) = el.ID()
			propID(num2) = p.ID()
			propname(num2) = p.title()
			grpID(num2) = grp.ID()
			grpname(num2) = grp.title()
			lcry(num2) = tmp(0)
			lcrz(num2) = tmp(1)
			lcrlt(num2) = tmp(2)
			length(num2) = tmp(3)
		Next
	Next

	App.feAppManagePanes ("Data Table", 1)
	dt.Lock( False )
	dt.Clear()

	'rc = dt.AddColumn(False, False, FT_ELEM, 0, "Element ID", FCT_INT, num + 1, elID, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Group ID", FCT_INT, num + 1, elID, grpID, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Group Name", FCT_STRING, num + 1, elID, grpname, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Property ID", FCT_INT, num + 1, elID, propID, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Property Name", FCT_STRING, num + 1, elID, propname, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Beam Length", FCT_DOUBLE, num + 1, elID, length, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Critical length 1", FCT_DOUBLE, num + 1, elID, lcry, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Critical length 2", FCT_DOUBLE, num + 1, elID, lcrz, num2)
	rc = dt.AddColumn(False, False, FT_ELEM, 0, "Critical length LT", FCT_DOUBLE, num + 1, elID, lcrlt, num2)

ExitHere:
	On Error Resume Next
	Set s = Nothing
	Set p = Nothing
	Set el = Nothing
	Set dt = Nothing
	Set grp = Nothing
	Set beams = Nothing
	Set App = Nothing
End Sub
