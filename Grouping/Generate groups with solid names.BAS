Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim g As femap.Group
    Set g = App.feGroup

    Dim s As femap.Solid
    Set s = App.feSolid

    Dim sset As femap.Set
    Set sset = App.feSet

    Dim sset1 As femap.Set
    Set sset1 = App.feSet

    Dim stitle As String

    sset.Select (FT_SOLID, True, "Select Solids to Generate Groups")

    sid = sset.First

    While sid > 0

    	sset1.Add (sid)

    	gid = g.NextEmptyID

    	App.feGroupGenSolid (sset1.ID)

    	g.Get (gid)

    	s.Get (sid)

    	stitle = s.title

    	g.title = stitle

    	g.Put (gid)

    	sset1.Clear

    	sid = sset.Next

	Wend

    App.feWindowRedraw	(App.Info_ActiveID (FT_VIEW))

End Sub
