Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim OutVec As femap.Output
    Set OutVec = App.feOutput

    Dim OutVec1 As femap.Output
    Set OutVec1 = App.feOutput

    Dim OutVec2 As femap.Output
    Set OutVec2 = App.feOutput

    Dim OutSet As femap.OutputSet
    Set OutSet = App.feOutputSet

    Dim Oset As femap.Set
    Set Oset = App.feSet

    Dim e As femap.Elem
    Set e = App.feElem

    Dim p As femap.Prop
    Set p = App.feProp

    Dim eset As femap.Set
    Set eset = App.feSet

    Dim m As femap.Matl
    Set m = App.feMatl

    Dim  EOutputSetID As Long

	rc = Oset.SelectID (FT_OUT_CASE, "Select Output Set", EOutputSetID)

	If rc = 2 Then
		GoTo Bailout
	End If

    OutSetID = OutSet.NextEmptyID

    OutSet.title = "Factor of Safety - From Output Set "+Str$(EOutputSetID)+"."

    OutSet.Put (OutSetID)

	eset.AddRule ( 25, FGD_ELEM_BYTYPE )
	eset.AddRule ( 26, FGD_ELEM_BYTYPE )

	OutVec1.InitScalarAtElem (OutSetID, 9000000,"Factor of Safety - Compression", FOT_STRESS, False)

	OutVec2.InitScalarAtElem (OutSetID, 9000001,"Factor of Safety - Tension", FOT_STRESS, False)

	App.Info_ActiveID (FT_OUT_DIR) = OutSetID

	OutSet.Get (EOutputSetID)
	OutVec.setID = EOutputSetID

	OutVec.Get (60031)

	eid = eset.First

	While eid > 0

		e.Get ( eid )

		outval = OutVec.Value (eid)

		PID = e.propID

		MatId = p.matlID

		m.Get (MatId)

		mtype = m.type

		Comp = m.mval (54)

		Tension = m.mval (52)

		e.Put ( eid )

		FoSC = Comp/outval

		OutVec1.setID = OutSetID

		OutVec1.ID = 9000000

		OutVec1.Value (eid) = FoSC

		FoST = Tension/outval

		OutVec2.setID = OutSetID

		OutVec2.ID = 9000001

		OutVec2.Value (eid) = FoST

		SkipFOS:
		eid = eset.Next

	Wend

	OutSet.ID = OutSetID

	OutVec1.Put (9000000)

	OutVec2.Put (9000001)

	Bailout:

	App.feViewRegenerate ( 0 )
    
End Sub
