Rem File: ElementCountByType.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim elSet2 As femap.Set
	Set elSet2 = App.feSet

	Dim elSet3 As femap.Set
	Set elSet3 = App.feSet

	Dim Msg As String

	Dim gset As femap.Set
	Set gset = App.feSet

	Dim gr As femap.Group
	Set gr = App.feGroup

	gset.SelectMultiID (FT_GROUP, 1, "Select Group(s)")

	Msg = "Model Element Summary"

	App.feAppMessage( FCL_BLACK, Msg )

	Dim j As Long

	While gset.Next > 0
		gid = gset.CurrentID

		Msg = "For Group"+Str$(gid)+":"

		App.feAppMessage( FCL_BLACK, Msg )

		For j=1 To 38

			rc = elSet.Clear()

			rc = elSet2.Clear()

			rc = elSet3.Clear()

			rc = elSet.AddRule( j, FGD_ELEM_BYTYPE )

			rc =elSet2.AddGroup (FT_ELEM, gid)

			rc = elSet3.AddCommon (elSet.ID, elSet2.ID)

			If elSet3.Count() > 0 Then

			Msg = "There are " + Str$(elSet3.Count())

			Select Case j
			Case 1
				Msg = Msg + " L_ROD elements"
			Case 2
				Msg = Msg + " L_BAR elements"
			Case 3
				Msg = Msg + " L_TUBE elements"
			Case 4
				Msg = Msg + " L_LINK elements"
			Case 5
				Msg = Msg + " L_BEAM elements"
			Case 6
				Msg = Msg + " L_SPRING elements"
			Case 7
				Msg = Msg + " L_DOF_SPRING elements"
			Case 8
				Msg = Msg + " L_CURVED_BEAM elements"
			Case 9
				Msg = Msg + " L_GAP elements"
			Case 10
				Msg = Msg + " L_PLOT elements"
			Case 11
				Msg = Msg + " L_SHEAR elements"
			Case 12
				Msg = Msg + " P_SHEAR elements"
			Case 13
				Msg = Msg + " L_MEMBRANE elements"
			Case 14
				Msg = Msg + " P_MEMBRANE elements"
			Case 15
				Msg = Msg + " L_BENDING elements"
			Case 16
				Msg = Msg + " P_BENDING elements"
			Case 17
				Msg = Msg + " L_PLATE elements"
			Case 18
				Msg = Msg + " P_PLATE elements"
			Case 19
				Msg = Msg + " L_PLANE_STRAIN elements"
			Case 20
				Msg = Msg + " P_PLANE_STRAIN elements"
			Case 21
				Msg = Msg + " L_LAMINATE_PLATE elements"
			Case 22
				Msg = Msg + " P_LAMINATE_PLATE elements"
			Case 23
				Msg = Msg + " L_AXISYM elements"
			Case 24
				Msg = Msg + " P_AXISYM elements"
			Case 25
				Msg = Msg + " L_SOLID elements"
			Case 26
				Msg = Msg + " P_SOLID elements"
			Case 27
				Msg = Msg + " L_MASS elements"
			Case 28
				Msg = Msg + " L_MASS_MATRIX elements"
			Case 29
				Msg = Msg + " L_RIGID elements"
			Case 30
				Msg = Msg + " L_STIFF_MATRIX elements"
			Case 31
				Msg = Msg + " L_CURVED_TUBE elements"
			Case 32
				Msg = Msg + " L_PLOT_PLATE elements"
			Case 33
				Msg = Msg + " L_SLIDE_LINE elements"
			Case 34
				Msg = Msg + " L_CONTACT elements"
			Case 35
				Msg = Msg + " L_AXISYM_SHELL elements"
			Case 36
				Msg = Msg + " P_AXISYM_SHELL elements"
			Case 37
				Msg = Msg + " P_BEAM elements"
			Case 38
				Msg = Msg + " L_WELD elements"
			End Select

		App.feAppMessage( FCL_BLACK, Msg )

		End If

	Next j

Wend

Msg = " "
App.feAppMessage( FCL_BLACK, Msg )

	
End Sub
