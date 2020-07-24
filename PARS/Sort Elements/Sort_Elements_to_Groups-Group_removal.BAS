Attribute VB_Name = "Module1"
Option Explicit

Const DBUG As Boolean = True
Private beams As Object

Sub Main()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long
	Dim i As Long
	Dim s As Object


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
					rc = -1
					Exit For
				End If
			End If
		End If
		rc = 0
	Next

	If rc <> -1 Then GoTo ExitHere

	Set s = App.feSet()

	Dim txt As Object
	Set txt = App.feText()

	rc = s.Select(FT_GROUP, True, "Select beam to delete:")
	If rc <> -1 Then GoTo ExitHere

	rc = s.Reset()
	For i = 0 To s.Count() - 1 Step 1
		rc = grp.Get(s.Next())
		rc = txt.Get(grp.ID)
		rc = txt.Delete(grp.ID)
		rc = beams.Remove(grp.ID)
		rc = grp.Delete(grp.ID)
	Next

	If beams.Count() < 1 Then
		beams.Save = False
	End If

ExitHere:
	On Error Resume Next
	Set txt = Nothing
	Set grp = Nothing
	Set s = Nothing
	Set beams = Nothing
	Set App = Nothing
End Sub
