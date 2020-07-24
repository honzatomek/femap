Rem File: SolidAlignCSys.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim cSysSET As femap.Set
	Set cSysSET = App.feSet
	Dim cSysID1 As Long
	Dim cSysID2 As Long

	Dim baseLoc(3) As Double
	Dim vbaseLoc As Variant
	Dim tipLoc(3) As Double
	Dim vtipLoc As Variant

	Dim moveVec(3) As Double
	Dim vmoveVec As Variant

	Dim solSET As femap.Set
	Set solSET = App.feSet
	Dim dist As Double

	Dim alVec1 As Variant
	Dim alVec2 As Variant

	rc = cSysSET.SelectID( FT_CSYS, "Select CSys 1 (solid to transform)", cSysID1 )
	rc = cSysSET.SelectID( FT_CSYS, "Select CSys 2 (reference)", cSysID2 )

	rc = solSET.Select( FT_SOLID, True, "Select Solids to Align to Reference)")

	'Figure out the base

	baseLoc(0) = 0.0
	baseLoc(1) = 0.0
	baseLoc(2) = 0.0

	vbaseLoc = baseLoc

	rc = App.feCoordTransform( cSysID1, vbaseLoc, 0, vbaseLoc )

	tipLoc(0) = 0.0
	tipLoc(1) = 0.0
	tipLoc(2) = 0.0

	vtipLoc = tipLoc

	rc = App.feCoordTransform( cSysID2, vtipLoc, 0, vtipLoc )

	moveVec(0) = vtipLoc(0) - vbaseLoc(0)
	moveVec(1) = vtipLoc(1) - vbaseLoc(1)
	moveVec(2) = vtipLoc(2) - vbaseLoc(2)

	vmoveVec = moveVec

	rc = App.feVectorLength( vmoveVec, dist )

	rc = App.feVectorUnit( vmoveVec, vmoveVec )

	rc = App.feMoveBy( FT_SOLID, solSET.ID, False, dist, vmoveVec )
	rc = App.feMoveBy( FT_CSYS, -cSysID1, False, dist, vmoveVec )

	'Everything is moved, now align x
	rc = App.feVectorAxis( cSysID1, 0, True, alVec1 )
	rc = App.feVectorAxis( cSysID2, 0, True, alVec2 )
	rc = App.feAlignTo( FT_SOLID, solSET.ID, vtipLoc, alVec1, vtipLoc, alVec2 )
	rc = App.feAlignTo( FT_CSYS, -cSysID1, vtipLoc, alVec1, vtipLoc, alVec2 )

	'Everything is moved, now align y
	rc = App.feVectorAxis( cSysID1, 1, True, alVec1 )
	rc = App.feVectorAxis( cSysID2, 1, True, alVec2 )
	rc = App.feAlignTo( FT_SOLID, solSET.ID, vtipLoc, alVec1, vtipLoc, alVec2 )
	rc = App.feAlignTo( FT_CSYS, -cSysID1, vtipLoc, alVec1, vtipLoc, alVec2 )

	'Everything is moved, now align x
	rc = App.feVectorAxis( cSysID1, 2, True, alVec1 )
	rc = App.feVectorAxis( cSysID2, 2, True, alVec2 )
	rc = App.feAlignTo( FT_SOLID, solSET.ID, vtipLoc, alVec1, vtipLoc, alVec2 )
	rc = App.feAlignTo( FT_CSYS, -cSysID1, vtipLoc, alVec1, vtipLoc, alVec2 )


	rc = App.feViewRegenerate( 0 )

	
End Sub
