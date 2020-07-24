Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim g As femap.Group
    Set g = App.feGroup

    Dim gset As femap.Set
    Set gset = App.feSet

    Dim gset2 As femap.Set
    Set gset2 = App.feSet

    Dim v As femap.View
    Set v = App.feView

    Dim viewID As Long

    Dim gmulti As Variant

    Dim gcount As Long

    Dim newgID As Long

    rc = App.feAppGetActiveView (viewID)

    v.Get(viewID)

    gstate = v.Group

    If gstate = 0 Then
    	rc = App.feAppMessageBox (0, "Multiple groups are not being displayed.  Choose 'Show Multiple' Groups to use this command")
    	GoTo Done
    End If

	If gstate = -1 Then
    	rc = App.feAppMessageBox (0, "Multiple groups are not being displayed.  Choose 'Show Multiple' Groups to use this command")
    	GoTo Done
    End If

    v.GetMultiGroupList (gcount, gmulti)

    If gcount = 0 Then
		rc = App.feAppMessage (FCM_ERROR, "No groups are currently being displayed.  No new group will be created")
		GoTo done
    End If

    If gcount = 1 Then
    	rc = App.feAppMessage (FCM_ERROR, "Only one group is being displayed.  New group would be redundant.")
		GoTo done
    End If

    For i = 0 To (gcount-1)

    	gID = gmulti(i)

    	If gid > 0 Then
    		gset.Add (gid)
    	End If

    	If gid < 0 Then
    		gid = -gid
    		gset2.Add (gid)
    	End If

    Next

    If gset.Count = 0 Then
    	rc = App.feAppMessage (FCM_ERROR, "Groups only removed from display.  No new groups created.")
		GoTo done
    End If

    If gset.Count <> 1 Then

    	rc = App.feGroupCombine (gset.ID, newgID)

    	g.Get(newgID)

    	g.title = "Group from Display"

    	g.Put(newgID)

    Else
    	gid = gset.First
		g.Get (gid)

		newgID = g.NextEmptyID

		g.title = "Group from Display"

		g.Put (newgID)
    End If

    If gset2.Count > 0 Then

    	subgroup = g.NextEmptyID

	    rc = App.feGroupBoolean2 (gset2.ID, newgID, FGB_SUBTRACT, 0, "Group from Display")

	    g.Get (subgroup)

	    g.Put (newgID)

	    g.Delete (subgroup)

	    App.Info_NextID (FT_GROUP) = subgroup

    End If

    Done:
    
End Sub
