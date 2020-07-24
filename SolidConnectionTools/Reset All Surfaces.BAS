Rem File: ResetAllSurfaces.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSet As femap.Set
	Dim feSurf As femap.Surface

	Set feSet = App.feSet()
	Set feSurf = App.feSurface()

	App.Info_LayerNoPick = 9999

	rc = feSurf.Reset

	While feSurf.Next
		feSurf.Layer = App.Info_ActiveID(FT_LAYER)
		rc = feSurf.Put(feSurf.ID)
		rc = feSet.Add(feSurf.ID)
	Wend

	rc = App.feModifyTransparency (FT_SURFACE,feSet.ID, 0.0 )

	rc = App.feViewRegenerate( 0 )

Jumping_Out:
End Sub
