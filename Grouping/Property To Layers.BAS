Rem File: PropertyToLayers.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

		Dim rc As Object
		Dim layerID As Integer
		Dim propID As Integer
		Dim elemSET As Object
		Dim aProp As Object
		Dim aLayer As Object

		Set aProp = App.feProp
		Set elemSET = App.feSet
		Set aLayer = App.feLayer

		propID = aProp.First()
		While propID <> 0

			'Get a new empty Layer
			layerID = aLayer.NextEmptyID
			aLayer.ID = layerID
			aLayer.title = aProp.title
			aLayer.color = aProp.color
			aLayer.Put(layerID)

			'Create a Set with all the elements of this property
			elemSET.Clear()
			elemSET.AddRule(aProp.ID, FGD_ELEM_BYPROP)

			' Update the layer reference for all these elements
			App.feModifyLayer ( FT_ELEM, elemSET.ID, layerID )

			propID = aProp.Next()
		Wend

End Sub
