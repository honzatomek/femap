Rem File: MakeSolidTransparent.BAS
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

	rc = feSet.SelectID(FT_SOLID, "Select Master Solid", masterSolidID )
	If masterSolidID = 0 Then
		GoTo Jumping_Out
	End If

	rc = msSurfSet.AddRule( masterSolidID, FGD_SURFACE_ONSOLID )
	rc = App.feModifyTransparency (FT_SURFACE, msSurfSet.ID, 70.0 )
	rc = App.feViewRegenerate( 0 )

Jumping_Out:
End Sub
