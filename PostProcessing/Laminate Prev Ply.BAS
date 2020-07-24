Sub Main
	Dim App As femap.model
    Set App = feFemap()
	Dim viewID As Long
	Dim feView As femap.View
	Set feView = App.feView

	rc = App.feAppGetActiveView( viewID )
	rc = feView.Get( viewID )
	feView.ContourData = feView.ContourData-200
	rc = feView.Put( viewID )
	rc = App.feViewRegenerate( 0 )
End Sub
