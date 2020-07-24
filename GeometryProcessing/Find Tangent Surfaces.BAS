Rem File: FindTangentSurfaces.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim suID As Long
	Dim suID1 As Long
	Dim suID2 As Long
	Dim suTEMP As Long
	Dim vAngle As Double

	Dim n1(3) As Double
	Dim n2(3) As Double
	Dim nv1 As Variant
	Dim nv2 As Variant

	Dim mdPoint(3) As Double
	Dim midV As Variant

	Dim inSet As femap.Set
	Dim tanSet As femap.Set
	Dim outSet As femap.Set
	Dim suWorkSet As femap.Set
	Dim cuSet As femap.Set
	Dim feCurve As femap.Curve
	Dim feSurf As femap.Surface

	Set inSet = App.feSet
	Set tanSet = App.feSet
	Set outSet = App.feSet
	Set cuSet = App.feSet
	Set suWorkSet = App.feSet
	Set feCurve = App.feCurve
	Set feSurf = App.feSurface

	midV = MidPoint
	nv1 = n1
	nv2 = n2

	rc = inSet.SelectID(5 , "Select Starting Surface",suID)

	rc = App.feAppLock

	rc = tanSet.Add( suID )

	If suID > 0 Then

	rc = inSet.Reset
		suID = inSet.Next

		While suID > 0

			'Add all the curves on this surface
			rc = cuSet.Clear
			rc = cuSet.AddRule( suID, 8 )

			If cuSet.count > 0 Then
				'Got curves to work on
				rc = cuSet.Reset
				cuID = cuSet.Next

				While  cuID > 0

					rc = feCurve.Get(cuID)
					rc = App.feCoordCurveMidpoint( cuID, midV )

					rc = suWorkSet.Clear
					rc = suWorkSet.AddRule( cuID, 10 )

					'remove any surfaces that have already been processed
					rc = tanSet.Reset
					suTEMP = tanSet.Next

					While suTEMP > 0
						If suTEMP <> suID Then
							rc = suWorkSet.Remove( suTEMP )
						End If
						suTEMP = tanSet.Next
					Wend

					If suWorkSet.count = 2 Then
						rc = suWorkSet.Reset
						suID1 = suWorkSet.Next
						suID2 = suWorkSet.Next
						'see which one is in the inSet
						If inSet.IsAdded( suID1 ) Then
							suID1 = suID1
							suID2 = suID2
						Else
							suTEMP = suID2
							suID2 = suID1
							suID1 = suTEMP
						End If
						'suID1 has the surface that is currently in
						'make sure suID2 has not already been eliminated
						'Check and see if the surfaces are tangent

						rc = App.feVectorNormalToSurface( suID1, midV, nv1 )
						rc = App.feVectorNormalToSurface( suID2, midV, nv2 )
						rc = App.feVectorAngle( nv1, nv2, vAngle )

						If vAngle < 0.1 Then
							rc = inSet.Add( suID2 )
							rc = tanSet.Add(suID2)
						End If
					End If

					cuID = cuSet.Next
				Wend
			End If

			rc = inSet.Remove( suID )
			rc = inSet.Reset
			suID = inSet.Next
		Wend
	End If
	rc = tanSet.Reset
	suID = tanSet.Next
	While suID > 0
		rc = feSurf.Get(suID)
		feSurf.color = 4
		rc = feSurf.Put(suID)
		suID = tanSet.Next
	Wend

	rc = App.feAppUnlock

	rc = App.feViewRegenerate( 0 )

End Sub
