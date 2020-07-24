Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim rigidSet As femap.Set
	Set rigidSet = App.feSet

	Dim elem1 As femap.Elem
	Set elem1 = App.feElem
	Dim elem2 As femap.Elem
	Set elem2 = App.feElem

    Dim color2 As Long

    Dim color3 As Long

    rc = rigidSet.AddRule( femap.FET_L_RIGID, FGD_ELEM_BYTYPE )

	i = rigidSet.First

    'rc = App.feAppColorPalette (FCL_WHITE, color2)

    'rc = App.feAppColorPalette (FCL_BLACK, color3)

    While i > 0

		If i = 0 Then
			GoTo Done
		End If

		rc = elem1.Get( i )

		If elem1.RigidInterpolate = True Then

			GoTo RBES
		End If

		'sets color of RBE2s to Crimson Red

		elem1.color = FCL_CRIMSON

        elem1.Put (i)

        GoTo AGAIN

        RBES:

        'sets color of RBE3s to Blue

        elem1.color = FCL_BLUE

		elem1.Put (i)

        AGAIN:

        i = rigidSet.Next
	Wend

    rc = App.feWindowRegenerate (1)

        DONE:

End Sub
