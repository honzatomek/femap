Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim el As femap.Elem
    Set el = App.feElem

    Dim elSet As femap.Set
    Set elSet= App.feSet
	Dim elNO As femap.Set
    Set elNO = App.feSet
	Dim gr As femap.Group
    Set gr = App.feGroup

	rc = elSet.Select(FT_ELEM, True, "Select Shells To Check for Material Angle" )

	If rc = 2 Then
		GoTo Done
	End If

	While elSet.Next > 0
		elID = elSet.CurrentID
		el.Get( elID )

		If el.topology = FTO_QUAD4 Or el.topology = FTO_QUAD8 Or el.topology = FTO_TRIA3 Or el.topology = FTO_TRIA6 Then
	    	If el.MaterialAngleFlag <> True Then
				elNO.Add( elID )
    		End If
		End If
	Wend

    	If elNO.Count = 0 Then
    		App.feAppMessage (FCM_NORMAL, "All selected elements either have a defined Material Angle or are not appropriate topology for Material Angle")
			GoTo Done
		Else
			gr.SetAdd ( FT_ELEM, elNO.ID )
			gr.title = "Elements without Defined Material Angle"
			gID = App.Info_NextID (FT_GROUP)
    		gr.Put( gid )
			App.feAppMessage (FCM_NORMAL,LTrim$(Str$(elNO.Count))+" Elements have no Material Angle defined and were placed into Group"+Str$(gid))
    	End If
    Done:

End Sub
