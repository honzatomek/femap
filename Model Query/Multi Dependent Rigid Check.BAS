Rem File: MultiDependentRigidCheck2.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim rigidSet As femap.Set
	Set rigidSet = App.feSet

	Dim elem1 As femap.Elem
	Set elem1 = App.feElem
	Dim elem2 As femap.Elem
	Set elem2 = App.feElem

	Dim i As Long
	Dim j As Long
	Dim k As Long
	Dim l As Long
	Dim m As Long
	Dim n As Long

	Dim ndCount1 As Long
	Dim ndCount2 As Long
	Dim ndList1 As Variant
	Dim ndList2 As Variant
	Dim faceList1 As Variant
	Dim faceList2 As Variant
	Dim weightList1 As Variant
	Dim weightList2 As Variant
	Dim DOF1 As Variant
	Dim DOF2 As Variant
	Dim Ref1 As Long
	Dim Ref2 As Long
	'Dim Indy1 As Long
	'Dim Indy2 As Long
	'Dim ntalist1 As Variant
	'Dim ntalist2 As Variant

	rc = rigidSet.AddRule( femap.FET_L_RIGID, FGD_ELEM_BYTYPE )

	While rigidSet.count > 1
		rc= rigidSet.Reset
		i = rigidSet.First
		If i = 0 Then
			GoTo Done
		End If

		rc = elem1.Get( i )

		If elem1.RigidInterpolate = True Then

			GoTo RBES
		End If

		rc = elem1.GetNodeList( 0, ndCount1, ndList1, faceList1, weightList1, DOF1 )
		Indy1 = elem1.Node (0)

		j = rigidSet.Next

		While j > 0
			rc = elem2.Get(j)

			If elem2.RigidInterpolate = True Then

			Ref2 = elem2.Node(0)
			rc = elem2.GetNodeList( 0, ndCount2, ntalist2, faceList2, weightList2, DOF2 )

				For k = 0 To ndCount1-1
					If  ndList1(k) = Ref2 Then
							Msg = "Node ID" + Str$( Ref2) + " is the RBE3 Reference Node of element" + Str$(j) + " and a RBE2 dependent node of element" + Str$(i)
							rc = App.feAppMessage( FCL_BLACK, Msg )
					End If
				Next k

				'For m = 0 To ndCount2-1
					'If  ntalist2(m) = Indy1 Then
							'Msg = "Node ID" + Str$(Indy1) + " is the RBE2 Independent Node of element" + Str$(j) + " and a RBE3 node to average of element" + Str$(i)
							'rc = App.feAppMessage( FCL_BLACK, Msg )
					'End If
				'Next m

			End If

			rc = elem2.GetNodeList( 0, ndCount2, ndList2, faceList2, weightList2, DOF2 )
				For k = 0 To ndCount1-1
					For l = 0 To ndCount2-1
						If ndList1(k) = ndList2(l) Then
							Msg = "Node ID" + Str$( ndList1(k)) + " is a RBE2 Dependent Node in both elements" + Str$(i) + " and " + Str$(j)
							rc = App.feAppMessage( FCL_BLACK, Msg )
						End If
					Next l
				Next k

			j = rigidSet.Next
		Wend
		rc = rigidSet.Remove( i )

		GoTo AGAIN

		RBES:

		Ref1 = elem1.Node (0)
		rc = elem1.GetNodeList( 0, ndCount1, ntalist1, faceList1, weightList1, DOF1 )

		j = rigidSet.Next

		While j > 0
			rc = elem2.Get(j)

			If elem2.RigidInterpolate = True Then

			Ref2 = elem2.Node(0)

					If Ref1 = Ref2 Then
							Msg = "Node ID" + Str$(Ref1) + " is the RBE3 Reference node in both elements" + Str$(i) + " and " + Str$(j)
							rc = App.feAppMessage( FCL_BLACK, Msg )
					End If

			End If

			rc = elem2.GetNodeList( 0, ndCount2, ndList2, faceList2, weightList2, DOF2 )
			Indy2 = elem2.Node (0)

					For l = 0 To ndCount2-1
						If Ref1 = ndList2(l) Then
							Msg = "Node ID" + Str$(Ref1) + " is the RBE3 Reference Node of element" + Str$(j) + " and a RBE2 dependent node of element" + Str$(i) '+ " and will likely cause Nastran to fail without use of PARAM, AUTOMPC"
							rc = App.feAppMessage( FCL_BLACK, Msg )
						End If
					Next l

					'For n = 0 To ndCount1-1
						'If  ntalist1(n) = Indy2 Then
							'Msg = "Node ID" + Str$(Indy2) + " is the RBE2 Independent Node of element" + Str$(j) + " and a RBE3 node to average of element" + Str$(i)
							'rc = App.feAppMessage( FCL_BLACK, Msg )
						'End If
					'Next n

			j = rigidSet.Next

		Wend

		rc = rigidSet.Remove( i )

		AGAIN:

	Wend

Done:

Msg = "Multiple Dependent Rigid Check Complete"

rc = App.feAppMessage( FCL_BLACK, Msg )

End Sub
