Rem File: SolidFaceLoadingwithEquation.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim surfID As Long
	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim cgXYZ As Variant
	Dim tcgXYZ As Variant 'Transformed Face Centroid

	Dim ldSet As Long
	Dim activeCSys As Long
	Dim ldID As Long

	Dim H As Double
	Dim T As Double

	Dim elLoad As femap.LoadMesh
	Set elLoad = App.feLoadMesh

	ldSet = App.Info_ActiveID( FT_LOAD_DIR )
	activeCSys = App.Info_ActiveID( FT_CSYS)

	surfID = 0

	rc = surfSet.SelectID( FT_SURFACE, "Select Surface to Apply Load To", surfID )

	If surfID > 0 Then

		rc = feSurf.Get( surfID )

		Dim nFaces As Long
		Dim eID As Variant
		Dim fID As Variant
		rc = feSurf.ElementFaces( False, True, False, nFaces, eID, fID )

		For i=0 To nFaces -1

			rc = feElem.Get( eID(i) )
			rc = feElem.GetFaceCentroid( fID(i), cgXYZ )
			rc = App.feCoordTransform( 0, cgXYZ, activeCSys, tcgXYZ )

			'Modify Loading Equations as needed
			H = 10*tcgXYZ(0)*tcgXYZ(0)-10*tcgXYZ(0)+100
			T = 100*tcgXYZ(0)*tcgXYZ(0)-100*tcgXYZ(0)+250.0

			elLoad.meshID = eID(i)
			elLoad.dof(0) = fID(i)
			elLoad.type = FLT_ECONVECTION
			elLoad.temp = T
			elLoad.ConvectionCoeff = H
			elLoad.setID = ldSet
			n = elLoad.NextEmptyID
			rc = elLoad.Put( elLoad.NextEmptyID )
		Next i

	End If

End Sub
