Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feView As femap.View
	Set feView = App.feView

	Dim viewID As Long

	rc = App.feAppGetActiveView( viewID )

	rc = feView.Get( viewID )

	If feView.Mode = FVM_FREE Then
		feView.Mode = FVM_HIDE
	Else
		feView.Mode = FVM_FREE
	End If

	rc = feView.Put( viewID )

	rc = App.feViewRegenerate( 0 )

End Sub
