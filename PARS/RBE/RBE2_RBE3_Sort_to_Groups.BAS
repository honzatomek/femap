Option Explicit

Const DBUG As Boolean = True

Sub Main
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long

	On Error GoTo ExitHere

	Dim elset As Object
	Set elset = App.feSet()

	Dim g2 As Object
	Set g2 = App.feGroup()

	Dim g3 As Object
	Set g3 = App.feGroup()

	Dim gset2 As Object
	Set gset2 = App.feSet()

	Dim gset3 As Object
	Set gset3 = App.feSet()

	Dim el As Object
	Set el = App.feElem()

	Dim gID As Long

	rc = elset.AddRule(29, FGD_ELEM_BYTYPE)
	If (rc <> -1) Or (elset.Count() = 0) Then GoTo ExitHere

	rc = elset.Reset()
	For i = 0 To elset.Count() - 1 Step 1
	 	rc = el.Get(elset.Next())
		If el.RigidInterpolate Then
			rc = gset3.Add(el.ID)
		Else
			rc = gset2.Add(el.ID)
		End If
	Next

	gID = g2.NextEmptyID()
	g2.title = "RBE2"
	rc = g2.Put(gID)

	gID = g3.NextEmptyID()
	g3.title = "RBE3"
	rc = g3.Put(gID)

	rc = g2.SetAdd(FT_ELEM, gset2.ID)
	rc = g3.SetAdd(FT_ELEM, gset3.ID)

	rc = g2.Put(g2.ID)
	rc = g3.Put(g3.ID)

ExitHere:
	On Error Resume Next
	Set gset2 = Nothing
	Set gset3 = Nothing
	Set el = Nothing
	Set g3 = Nothing
	Set g2 = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub
