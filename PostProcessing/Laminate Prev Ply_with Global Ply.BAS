Sub Main
	Dim App As femap.model
    Set App = feFemap()
	Dim viewID As Long
	Dim feView As femap.View
	Set feView = App.feView

	Dim gp As femap.globalply
	Set gp = App.feGlobalPly

	Dim ov As femap.Output
	Set ov = App.feOutput

	rc = App.feAppGetActiveView( viewID )
	rc = feView.Get( viewID )
	outsetID = feView.OutputSet
	globalplyID = feView.GlobalPlyLocation
	If globalplyID > 0 Then
		gp.ID = globalplyID
		gpID = gp.PrevID
		If gpID > 0 Then
			feView.GlobalPlyLocation = gpID
		Else
			App.feAppMessage (FCM_WARNING, "That was the first Global Ply ID, use Laminate Next Ply")
		End If
	Else
		ovID = feView.ContourData -200
		ov.setID = outsetID
		If ov.Exist (ovID) = -1 Then
			feView.ContourData = feView.ContourData-200
		Else
			App.feAppMessage (FCM_WARNING, "Output Vector does not exist")
		End If
	End If
	rc = feView.Put( viewID )
	rc = App.feViewRegenerate( 0 )
End Sub
