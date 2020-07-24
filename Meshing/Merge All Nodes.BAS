Rem File: MergeAllNodes.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim listCoinc As Boolean
	Dim mergeSummary As Boolean
	Dim mergeTol As Double

	listCoinc = False
	mergeSummary = False
	mergeTol = .001

	rc = ndSet.AddAll( FT_NODE)

	Dim ndCount As Long
	Dim coinNodes As Variant

	App.feCheckCoincidentNode( ndSet.ID, ndSet.ID, mergeTol, True, listCoinc, ndCount, coinNodes )

	If (listCoinc = False And mergeSummary = True) Then
		Dim Msg As String
		Msg = " Node"
		If (ndCount <> 1) Then
		  Msg = Msg +"s"
		End If

		Msg = Str$(ndCount) + Msg + " merged"
		App.feAppMessage( FCL_BLACK, Msg )

   End If

	rc = App.feViewRegenerate( 0 )

End Sub
