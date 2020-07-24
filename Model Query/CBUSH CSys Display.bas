Sub Main
    Dim App As femap.model
    Set App = feFemap()

	Dim GFXA As femap.GFXArrow
    Set GFXA = App.feGFXArrow

    Dim ArrowSet As femap.Set
    Set ArrowSet = App.feSet

    Dim eset As femap.Set
	Set eset = App.feSet

	Dim Coord As femap.CSys
	Set Coord = App.feCSys

	Dim e As femap.Elem
	Set e = App.feElem

	Dim p As femap.Prop
	Set p = App.feProp

	Dim n As femap.Node
	Set n = App.feNode

	Dim v As femap.View
    Set v = App.feView

    Dim ViewID As Long
    Dim CoordOrigin As Variant
    Dim CoordVec (2) As Double
    Dim PoleVec (2) As Double
    Dim ThetaVec As Variant
    Dim PhiVec As Variant

    App.feAppGetActiveView (ViewID)

    v.Get (ViewID)

    CSysArrow = v.ColorMode (FVI_CSYS)

    Select Case CSysArrow
    Case 0, 1, 2, 3, 4, 5, 6, 7
    	ArrowType = 0
    Case 8, 9, 10, 11
    	ArrowType = 1
    End Select

	Dim DirCos As Variant

	Dim size As Double

	'App.feGetReal ("Enter Arrow Size", 1E-8, 10, size)

	size = 1.0

	'eset.AddAll (FT_ELEM)

	eset.AddRule (6, FGD_ELEM_BYTYPE)

    While eset.Next > 0
    	eid = eset.CurrentID
    	e.Get (eid)
			propID = e.propID
			p.Get (propID)
			isBush = p.flag (3)
			If isBush = True Then

				nid = e.Node (0)
				n.Get (nid)

				nx = n.x
				ny = n.y
				nz = n.z

				refcoord = p.refCS
				Coord.Get (refcoord)
				CoordType = Coord.type
				DirCos = Coord.matrix

			If CoordType <> 2 Then

				If CoordType = 0 Then

   				ArrowID1 =GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID1, nx, ny, nz, DirCos(0), DirCos(1), DirCos(2), size, GAM_SCALED, 1, 34, ArrowType)

    			ArrowSet.Add (ArrowID1)

    			ArrowID2 =GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID2, nx, ny, nz, DirCos(3), DirCos(4), DirCos(5), size, GAM_SCALED, 1, 72, ArrowType)

    			ArrowSet.Add (ArrowID2)

				ArrowID3 =GFXA.NextEmptyID

   				rc=GFXA.PutAll (ArrowID3, nx, ny, nz, DirCos(6), DirCos(7), DirCos(8), size, GAM_SCALED, 1, 106, ArrowType)

    			ArrowSet.Add (ArrowID3)

    			Else

				CoordOrigin = Coord.origin

				CoordVec(0) = nx - CoordOrigin(0)
				CoordVec(1) = ny - CoordOrigin(1)
				CoordVec(2) = nz - CoordOrigin(2)

				PoleVec(0) = DirCos(6)
				PoleVec(1) = DirCos(7)
				PoleVec(2) = DirCos(8)

				App.feVectorCrossProduct (PoleVec,CoordVec, PhiVec)
				App.feVectorCrossProduct (PhiVec, PoleVec, ThetaVec)

				ArrowID7 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID7, nx, ny, nz, ThetaVec(0), ThetaVec(1), ThetaVec(2), size, GAM_SCALED, 1, 24, ArrowType)

				ArrowSet.Add (ArrowID7)

				ArrowID8 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID8, nx, ny, nz, PhiVec(0), PhiVec(1), PhiVec(2), size, GAM_SCALED, 1, 14, ArrowType)

				ArrowSet.Add (ArrowID8)

				ArrowID9 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID9, nx, ny, nz, PoleVec(0), PoleVec(1), PoleVec(2), size, GAM_SCALED, 1, 120, ArrowType)

				ArrowSet.Add (ArrowID9)

				End If

			Else
				CoordOrigin = Coord.origin

				CoordVec(0) = nx - CoordOrigin(0)
				CoordVec(1) = ny - CoordOrigin(1)
				CoordVec(2) = nz - CoordOrigin(2)

    			ArrowID4 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID4, nx, ny, nz, CoordVec(0), CoordVec(1), CoordVec(2), size, GAM_SCALED, 1, 82, ArrowType)

				ArrowSet.Add (ArrowID4)

				PoleVec(0) = DirCos(6)
				PoleVec(1) = DirCos(7)
				PoleVec(2) = DirCos(8)

				'App.feVectorCrossProduct (CoordVec, PoleVec, PhiVec)
				App.feVectorCrossProduct (PoleVec, CoordVec, PhiVec)
				'App.feVectorCrossProduct (CoordVec, PhiVec, ThetaVec)
				App.feVectorCrossProduct (PhiVec, CoordVec, ThetaVec)

				ArrowID5 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID5, nx, ny, nz, ThetaVec(0), ThetaVec(1), ThetaVec(2), size, GAM_SCALED, 1, 7, ArrowType)

				ArrowSet.Add (ArrowID5)

				ArrowID6 = GFXA.NextEmptyID

    			rc=GFXA.PutAll (ArrowID6, nx, ny, nz, PhiVec(0), PhiVec(1), PhiVec(2), size, GAM_SCALED, 1, 132, ArrowType)

				ArrowSet.Add (ArrowID6)

			End If

    	End If

    Wend

	rc = App.feGFXSelect (ArrowSet.ID, True, True)

	App.feAppMessage (FCM_NORMAL, "Color Key")
	App.feAppMessage (FCM_NORMAL, "Rectangular - X Direction = Red, Y Direction = Green, Z Direction = Blue")
	App.feAppMessage (FCM_NORMAL, "Cylindrical - Radial Direction = Yellow, Theta Direction = Orange, Z Direction = Cyan")
	App.feAppMessage (FCM_NORMAL, "Spherical - Radial Direction = Purple, Theta Direction = Brown, Phi Direction = Gray")
    
End Sub
