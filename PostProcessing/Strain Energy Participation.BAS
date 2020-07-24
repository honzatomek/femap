Sub Main
    Dim App As femap.model
	Set App = GetObject(,"femap.model")

	Dim g As femap.Group
	Set g = App.feGroup

	Dim gset As femap.Set
	Set gset = App.feSet

	Dim glist As Set

	Dim out As femap.Output
	Set out = App.feOutput

	Dim outset As femap.OutputSet
	Set outset = App.feOutputSet

	Dim oset As femap.Set
	Set oset = App.feSet

	Dim value As Variant
	Dim gID As Long

	oset.Select (FT_OUT_CASE, True, "Select Output Sets")

	gset.SelectMultiID (FT_GROUP, 1, "Select Group(s)")

	gID = gset.First

	osetID = oset.First

	While osetID > 0

		outset.Get(osetID)

		outval = outset.Value

		App.feAppMessage (FCM_NORMAL, "Output Set " + osetID + ", Mode " + outval + "")

		out.GetFromSet (osetID, 80000)

        totoutval = (((outval)*(2)*(3.14159265358979))^2)/2

		While gID > 0

			g.Get (gID)

			grID = gID

			Set glist = g.List (FGR_ELEM)

			out.GetOutputListAtSet (glist.ID, 0.0, value)

			ecount = glist.Count

			i = 0
			total = 0

			For i = 0 To (ecount-1)
				eID = glist.Next
                strainval = value(i)
                If strainval > 0.0 Then
                	total = ((strainval/totoutval)*100) + total
				End If

			Next i

			App.feAppMessage (FCM_NORMAL, "Percent Strain Energy in group " + grID + " is " + Total + "")

		gID = gset.Next

		Wend

		gID = gset.First

		App.feAppMessage (FCM_NORMAL,"")

		osetID = oset.Next

	Wend
End Sub
