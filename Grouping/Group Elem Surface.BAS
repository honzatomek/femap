Rem File: GroupElemSurface.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim suSet As femap.Set
	Set suSet = App.feSet

	Dim feGroup As femap.Group
	Set feGroup = App.feGroup

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim elID As Long

	grID = App.Info_ActiveID( FT_GROUP )

	rc = feGroup.Get( grID )

	rc = elSet.SelectID( FT_ELEM, "Select Element", elID )

	If elID > 0 Then
		rc = feElem.Get( elID )
		If feElem.geomID > 0 Then
			rc = feGroup.RangeAdd( FGD_SURFACE_ID, feElem.geomID , 0, 1, 1 )
		End If
	End If



	rc = feGroup.Put( grID )

	rc = App.feViewRegenerate( 0 )

	
End Sub
