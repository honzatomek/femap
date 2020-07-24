'Renumbers RBE Elements with the number of Independent node

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim el As Object
    Set el = App.feElem()
    Dim eset As Object
    Set eset = App.feSet()
    Dim eID As Long
    Dim ndID As Long
    Dim i As Long
    Dim ecount As Long

    rc = eset.Select( 8, True, "Select Rigid Elements to be renumbered by master node")
    If rc <> - 1 Then Exit Sub
	ecount = eset.Count()
	rc = eset.Reset()

	For i = 0 To ecount - 1
		eID = eset.Next()
		rc = el.Get(eID)
		If (el.type = 29) And (el.topology = 13) Then
			nID = el.Node(0)
			rc = App.feRenumberOpt2(FT_ELEM, -eID, nID, 0, 0, False, False, False, 0)
		End If
	Next

    Set el = Nothing
    Set eset = Nothing
    Set App = Nothing
End Sub
