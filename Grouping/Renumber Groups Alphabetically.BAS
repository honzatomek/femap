Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim g1 As femap.Group
    Set g1 = App.feGroup

	Dim g2 As femap.Group
    Set g2 = App.feGroup

    Dim gset As femap.Set
    Set gset = App.feSet

    If App.Info_Count(FT_GROUP) = 0 Then
		Msg = "No Groups in Model, Exiting..."
		rc = MsgBox( Msg, vbOkOnly )
		GoTo OK
	End If

	Dim gcount As Long
    Dim listID As Variant
    Dim gtitles As Variant

    g1.GetTitleList (0, 0, gcount, listID, gtitles)

'Bubble Sort

For i = gcount-1 To 0 Step -1

       	For j = 1 To i

            title1 = gtitles(j)

            title2 = gtitles(j-1)

            value = StrComp (title1, title2, vbTextCompare)

                If Value = -1 Then

					temp = listID (j-1)

                    temptitle = gtitles (j-1)

                    listID (j-1) = listID (j)

                    gtitles (j-1) = gtitles (j)

					listID (j) = temp

                    gtitles (j) = temptitle

                End If

        Next j

Next i

For i = 0 To gcount -1

g1.Get (listID(i))

g1.Put (i + 90000000)

g1.Delete (listID(i))

Next i

gset.AddAll (FT_GROUP)

App.feRenumberOpt (FT_GROUP, 1, 1, 0, 0, False, False, False)

OK:

End Sub
