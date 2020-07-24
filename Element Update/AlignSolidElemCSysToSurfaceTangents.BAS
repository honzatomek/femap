Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim elemSET As femap.Set
    Set elemSET = App.feSet

    rc = elemSET.Select( FT_ELEM, True, "Select Elements to Update Alignment" )

    Dim feProp As femap.Prop
    Set feProp = App.feProp

    Dim feSurf As femap.Surface
    Set feSurf = App.feSurface

    Dim feElem As femap.Elem
    Set feElem = App.feElem

    rc = feSurf.SelectID( "Select Surface to Align With" )

    Dim xyz As Variant
    Dim xyz_x(2) As Double
    Dim xyz_y(2) As Double
  	Dim u As Double
  	Dim v As Double

  	Dim sTan As Variant
  	Dim vTan As Variant

  	Dim feCSys As femap.CSys
  	Set feCSys = App.feCSys
  	Dim propID As Long

    Dim elID As Long

    elID = elemSET.First()

    While elID > 0
		rc = feElem.Get( elID )
		rc = feElem.GetCentroid( xyz )
		rc = feSurf.XYZToParam( xyz, u, v )
		rc = feSurf.Tangents( u, v, sTan, vTan )
		' Found the alignment vectors at this location, create a coordinate system
		xyz_x(0) = xyz(0) + sTan(0)
		xyz_x(1) = xyz(1) + sTan(1)
		xyz_x(2) = xyz(2) + sTan(2)
		xyz_y(0) = xyz(0) + vTan(0)
		xyz_y(1) = xyz(1) + vTan(1)
		xyz_y(2) = xyz(2) + vTan(2)
		rc = feCSys.XYPoints( xyz, xyz_x, xyz_y )
		rc = feCSys.Put( feCSys.NextEmptyID )
		If rc = -1 Then
			rc = feProp.Get( feElem.propID )
			feProp.refCS = feCSys.ID
			propID = feProp.NextEmptyID
			rc = feProp.Put( feProp.NextEmptyID )
			feElem.propID = propID
			rc = feElem.Put( feElem.ID )
		End If


    	elID = elemSET.Next()
    Wend

    End Sub
