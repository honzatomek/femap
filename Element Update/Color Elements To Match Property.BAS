Rem File: ColorElementsToMatchProperty.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim feProp As femap.Prop
	Set feProp = App.feProp

	Dim lastProp As Long

	lastProp = 0

	While feElem.Next
		If lastProp <> feElem.propID Then
			lastProp = feElem.propID
			rc = feProp.Get( feElem.propID)
		End If
		feElem.color = feProp.color
		feElem.Put( feElem.ID )
	Wend
	rc = App.feViewRegenerate( 0 )

End Sub
