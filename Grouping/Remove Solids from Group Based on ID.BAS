Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim Sol As femap.Solid
    Set Sol = App.feSolid

    Dim S As femap.Surface
    Set S = App.feSurface

    Dim C As femap.Curve
    Set C = App.feCurve

    Dim P As femap.Point
    Set P = App.fePoint

    Dim E As femap.Elem
    Set E = App.feElem

    Dim N As femap.Node
    Set N = App.feNode

    Dim Solset As femap.Set
    Set Solset = App.feSet
    
    Dim SSet As femap.Set
    Set SSet = App.feSet

    Dim CSet As femap.Set
    Set CSet = App.feSet

    Dim PSet As femap.Set
    Set PSet = App.feSet

    Dim ESet As femap.Set
    Set ESet = App.feSet

    Dim ESet2 As femap.Set
    Set ESet2 = App.feSet

    Dim NSet As femap.Set
    Set NSet = App.feSet

    Dim NSet2 As femap.Set
    Set NSet2 = App.feSet

    Dim g As femap.Group
    Set g = App.feGroup

    Solset.Select (FT_SOLID, True, "Select Solids to Remove From Group")

    SolidID = Solset.First

    While solidID > 0

    	SSet.AddRule (SolidID, FGD_SURFACE_ONSOLID)

    	CSet.AddSetRule (SSet.ID, FGD_CURVE_ONSURFACE)

    	PSet.AddSetRule (CSet.ID, FGD_POINT_ONCURVE)

        ESet.AddRule (SolidID, FGD_ELEM_ATSOLID)

        NSet.AddRule (SolidID, FGD_NODE_ATSOLID)

    	SolidID = Solset.Next

    Wend

	gid = App.Info_ActiveID (FT_GROUP)

    g.Get (gid)

    g.Condense = True

    g.Put (gid)

    g.Get (gid)

    g.SetAddOpt (FT_POINT, PSet.ID, 0)

    g.SetAddOpt (FT_CURVE, CSet.ID, 0)

    g.SetAddOpt (FT_SURFACE, SSet.ID, 0)

    g.SetAddOpt (FT_NODE, NSet.ID, 0)

    g.SetAddOpt (FT_ELEM, ESet.ID, 0)

    g.SetAddOpt (FT_SOLID, Solset.ID, 0)

    g.Put (gid)

	g.Get (gid)

    ESet2.AddGroup (FT_ELEM, gid)

    NSet2.AddSetRule (ESet2.ID, FGD_NODE_ONELEM)

    g.SetAdd (FT_NODE, NSet2.ID)

    g.Put (gid)

	rc = App.feWindowRegenerate (0)

End Sub
