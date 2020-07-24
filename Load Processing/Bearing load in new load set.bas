Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim sLoad As femap.LoadGeom
    Set sLoad = App.feLoadGeom

    Dim eload As femap.LoadMesh
    Set eload = App.feLoadMesh

    Dim sSet As femap.Set
    Set sSet = App.feSet

    Dim eset As femap.Set
    Set eset = App.feSet

    Dim ldset As femap.LoadSet
    Set ldset = App.feLoadSet

    Dim lddef As femap.LoadDefinition
    Set lddef = App.feLoadDefinition

    Dim cs As femap.CSys
	Set cs = App.feCSys

    Dim nCS As Long
    Dim csID As Long
	Dim vID As Variant
	Dim vTitle As Variant
	Dim CsysID As Long
    Dim vlist() As String

    cs.GetTitleIDList (True, 0, 0, nCS, vID, vTitle)

    ReDim vlist(2+nCS)
    'vlist(0) = "0..Basic Rectangular"
    vlist(0) = "1..Basic Cylindrical"
    'vlist(2) = "2..Basic Spherical"

 	For i=0 To nCS-1

        cs.Get (vID(i))

        If cs.type = FCS_CYLINDRICAL Then

	    vlist(1+i) = vTitle(i)

        End If
    Next

    Dim Csyslist$()

    Csyslist$() = vlist

    Dim ldSetID As Long

    'ldSetID = App.Info_ActiveID( FT_LOAD_DIR )

    If ldSetID = 0 Then

        ldSetID = ldset.NextEmptyID

        ldset.title = "Bearing Load Set"

        ldset.Put (ldSetID)

       App.Info_ActiveID( FT_LOAD_DIR ) = ldSetID

       sLoad.setID = ldSetID

    End If

	rc = sSet.Select (FT_SURFACE, True , "Select Surfaces to Apply Bearing Load")

	If rc = 2 Then
		GoTo Done

	End If

	Begin Dialog UserDialog 450,147,"Bearing Load Values" ' %GRID:10,7,1,1
		OKButton 90,105,120,28
		CancelButton 220,105,140,28
		text 10,21,180,14,"Choose Coordinate System",.Text1
		text 10,56,180,14,"Load Value",.Text2
		DropListBox 190,21,240,21,Csyslist(),.Csyslist,1
		TextBox 190,56,140,21,.lvalue
	End Dialog
	Dim dlg As UserDialog

    If Dialog(dlg) = 0 Then
    GoTo Done
	End If

    CsysID = Val(Split(dlg.Csyslist, "..", 2) (0) )

    sID = sSet.First

    ldID = lddef.NextEmptyID

    lddef.setID = ldSetID
    lddef.loadTYPE = FLT_SEPRESSURE
    lddef.DataType = FT_GEOM_LOAD
    lddef.title = "Bearing Load on Elements"
    lddef.Put (ldID)

    While sID > 0

	sLoad.type = FLT_SEPRESSURE
    sLoad.geomID = sID
    sLoad.CSys = CsysID
    sLoad.variation = FLV_EQUATION
    sLoad.varEqn = "sin(!y)"
    sLoad.Pressure = Val(dlg.lvalue)
    sLoad.LoadDefinitionID = ldID
    sLoad.Put (sLoad.NextEmptyID)

    sid = sSet.Next

    Wend

    ldset.Get (ldSetID)

    ldset.Expand ()

    eset.AddAll (FT_ELEM)

    eload.setID = ldSetID

    If eset.count() > 0 Then
            While eload.Next

            If eload.LoadDefinitionID = ldID Then

            	eload.expanded = False

                eload.Put(eload.ID)

	            If eload.Pressure < 0 Then

        		eload.Delete (0)

                End If

            End If
            Wend
    End If

    sLoad.Reset

    While sLoad.Next

		If sLoad.setID = ldSetID Then

    		If sLoad.LoadDefinitionID = ldID Then

    		sLoad.Delete (0)

    		End If

        End If
	Wend

    lddef.Get (ldID)
    lddef.loadTYPE = FLT_EPRESSURE
    lddef.DataType = FT_SURF_LOAD
    lddef.Put (ldID)

    ldset.Compress ()

    rc = App.feViewRegenerate( 0 )

    Done:

End Sub


