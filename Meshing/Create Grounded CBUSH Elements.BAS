'This API script will create a "coincident node" CBUSH Element at each selected node.
'The 2nd node of the CBUSH will be created and have all 6 DOF permanently constrained.
'A CBUSH property with an Orinetation CSys must be selected in order for this API to run to completion.

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim n As femap.Node
    Set n = App.feNode

    Dim nset As femap.Set
    Set nset = App.feSet

    Dim p As femap.Prop
    Set p = App.feProp

    Dim pset As femap.Set
    Set pset = App.feSet

    Dim e As femap.Elem
    Set e = App.feElem

    Dim pid As Long

    pset.AddAll (FT_PROP)

    If pset.Count = 0 Then
    	rc = App.feAppMessageBox (0, "No properties exist. Create a Spring/Damper Property and try again")
		GoTo Done
    End If

    Again:

    pset.SelectID (FT_PROP, "Select Spring/Damper Property", pid)

    p.Get (pid)
    ptype = p.type

    If ptype = 6 Then
    	If p.flag (0) = False Then
			rc = App.feAppMessageBox (2,"Spring/Dmaper Property does not have an Orientation Coordinate System. Choose Another Property? (No = Cancel)")
			If rc = -1 Then
				GoTo Again
			Else
				GoTo Done
			End If
    	End If

    Else
    	rc = App.feAppMessageBox (2, "Not a Spring/Dmaper Property. Choose Another Property? (No = Cancel)")
    	If rc = -1 Then
			GoTo Again
		Else
			GoTo Done
		End If
    End If

    nset.Select (FT_NODE, True, "Select Nodes for Grounded CBUSH Elements")

    While nset.Next
    	nid = nset.CurrentID
    	n.Get (nid)
    	nidG = n.NextEmptyID
    	n.permBC (0) = True
    	n.permBC (1) = True
    	n.permBC (2) = True
    	n.permBC (3) = True
    	n.permBC (4) = True
    	n.permBC (5) = True
    	n.Put (nidG)
		e.type = FET_L_SPRING
		e.topology = FTO_LINE2
		e.Node (0) = nid
		e.Node (1) = nidG
		e.propID = pid
		e.formulation (1) = 1
    	e.Put (e.NextEmptyID)
    Wend

    Done:
    
End Sub
