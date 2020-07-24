Rem File: IsolateSolid.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSet As femap.Set
	Dim msSurfSet As femap.Set
	Dim otherSurfSet As femap.Set
	Dim feSurf As femap.Surface

	Set feSet = App.feSet()
	Set msSurfSet = App.feSet()
	Set feSurf = App.feSurface()
	Set otherSurfSet = App.feSet()

	Dim masterSolidID As Long

	masterSolidID = 0

	App.Info_LayerNoPick = 0
	rc = App.feViewRegenerate( 0 )
	rc = feSet.SelectID(FT_SOLID, "Select Master Solid", masterSolidID )
	App.Info_LayerNoPick = 9999

	If masterSolidID = 0 Then
		GoTo Jumping_Out
	End If

	rc = msSurfSet.AddRule( masterSolidID, FGD_SURFACE_ONSOLID )

	If msSurfSet.count = 0 Then
		GoTo Jumping_Out
	End If

	rc = feSurf.Reset
	rc = otherSurfSet.Clear

	While feSurf.Next
		If msSurfSet.IsAdded ( feSurf.ID ) Then
			feSurf.Layer = App.Info_ActiveID( FT_LAYER )
			rc = feSurf.Put( surfID )
		Else
			feSurf.Layer = App.Info_LayerNoPick
			rc = otherSurfSet.Add ( feSurf.ID)
			rc = feSurf.Put( surfID )
		End If
	Wend

	rc = App.feModifyTransparency (FT_SURFACE, otherSurfSet.ID, 70.0 )
	rc = App.feModifyTransparency (FT_SURFACE, msSurfSet.ID, 0.0 )

	rc = App.feViewRegenerate( 0 )

Jumping_Out:
End Sub
