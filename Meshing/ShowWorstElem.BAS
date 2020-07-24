Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim viewID As Long
	Dim feView As femap.View
	Set feView = App.feView

	Dim grID As Long

	grID = App.Info_NextID( FT_GROUP )

	rc = elSet.AddAll(FT_ELEM)

	rc = App.feCheckElemDistortion( elSet.ID, False, True, False, False, False, False, False, False, False, True, False, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.7, 0.0)

	Dim elSet2 As femap.Set
	Set elSet2 = App.feSet

	rc = elSet2.AddGroup( FT_ELEM, grID )

	Dim feElem As femap.Elem
	Set feElem = App.feElem
	Dim elemID As Long

	Dim hJ As Double
	hJ = 0.0
	Dim worstElem As Long
	worstElem = 0

	Dim elemCenter As Variant

	Dim v1 As Double
	Dim v2 As Double
	Dim v3 As Double
	Dim v4 As Double
	Dim v5 As Double
	Dim v6 As Double
	Dim v7 As Double
	Dim v8 As Double
	Dim v9 As Double

	elemID = elSet2.First
	While elemID > 0
		App. feGetElemDistortion( elemID, v1, v2, v3, v4, v5, v6, v7, v8, v9)
		If v8 > hJ Then
			hJ = v8
			worstElem = elemID
		End If

		elemID = elSet2.Next
	Wend

	If worstElem > 0 Then
		Msg = "Worst Element ID = " + Str$( worstElem ) + " Jacobian = " + Str$(hJ)
		rc = App.feAppMessage( FCM_NORMAL, Msg )
		rc = feElem.Get( worstElem)
		rc = App.feAppGetActiveView( viewID )
		If rc = -1 Then
			rc = feView.Get( viewID )
			rc = feElem.GetCentroid( elemCenter )
			feView.center(0) = elemCenter(0)
			feView.center(1) = elemCenter(1)
			feView.center(2) = elemCenter(2)
			rc = feView.Put( viewID )
			rc = App.feViewRegenerate( viewID )
		End If
		'If feElem.geomID > 0 Then
		'	rc= App.feWindowShow( FT_SURFACE, -feElem.geomID )
		'Else
			rc = App.feWindowShow( FT_ELEM, -feElem.ID )
		'End If
	End If

End Sub
