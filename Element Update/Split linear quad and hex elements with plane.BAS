Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim e As femap.Elem
    Set e = App.feElem

    Dim eset As femap.Set
    Set eset = App.feSet

    Dim n As femap.Node
    Set n = App.feNode

    Dim nset As femap.Set
    Set nset = App.feSet

    Dim nsetp As femap.Set
    Set nsetp = App.feSet

    Dim vn1 As Variant
    Dim vn2 As Variant
    Dim vn3 As Variant
    Dim vn4 As Variant
	Dim vn5 As Variant
    Dim vn6 As Variant
    Dim vn7 As Variant
    Dim vn8 As Variant

	Dim PlaneO As Variant
    Dim PlaneN As Variant
    Dim PlaneA As Variant
    Dim PlaneP As Variant
    Dim PlaneC As Variant

    Dim dist As Double
    Dim edist As Double
    Dim edist1 As Double
    Dim edist2 As Double

    'Scenerio #1 vectors

    Dim vec14(2) As Double
    Dim vec23(2) As Double
    Dim vec58(2) As Double
    Dim vec67(2) As Double

    'Scenerio #2 vectors
    Dim vec12(2) As Double
    Dim vec34(2) As Double
    Dim vec56(2) As Double
    Dim vec78(2) As Double

    'Scenerio #3 vectors
    Dim vec15(2) As Double
    Dim vec26(2) As Double
    Dim vec37(2) As Double
    Dim vec48(2) As Double

    App.feAppMessage (FCM_WARNING, "If you would like to increase the efficiency of this API, select only elements which have a chance to be split by the specified plane")

    rc = eset.Select (FT_ELEM, True, "Select Elements to Split")

    If rc = 2 Then
    	GoTo Done
    End If

    rc = App.fePlanePick ("Select Plane to Slice Elements", PlaneO, PlaneN, PlaneA)

    If rc = 2 Then
    	GoTo Done
    End If

    App.feAppMessage (FCM_WARNING, "This API can be computationally intensive.  Please allow it time to run and look for a message that it has 'finished'")

    App.feAppLock

    eID = eset.First

    While eID > 0
    	e.Get (eID)

    	If e.topology = FTO_BRICK8 Then

			n1 = e.Node (0)
        	n2 = e.Node (1)
        	n3 = e.Node (2)
        	n4 = e.Node (3)
            n5 = e.Node (4)
        	n6 = e.Node (5)
        	n7 = e.Node (6)
        	n8 = e.Node (7)

            nset.Add(n1)
            nset.Add(n2)
            nset.Add(n3)
            nset.Add(n4)
            nset.Add(n5)
            nset.Add(n6)
            nset.Add(n7)
            nset.Add(n8)

        	n.Get (n1)
        	vn1 = n.xyz
            App.feMeasureDistanceToPlane (vn1, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n1+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

        	n.Get (n2)
        	vn2 = n.xyz
            App.feMeasureDistanceToPlane (vn2, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n2+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

        	n.Get (n3)
        	vn3 = n.xyz
            App.feMeasureDistanceToPlane (vn3, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n3+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

			n.Get (n4)
        	vn4 = n.xyz
			App.feMeasureDistanceToPlane (vn4, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n4+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

            n.Get (n5)
        	vn5 = n.xyz
			App.feMeasureDistanceToPlane (vn5, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n5 + " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

        	n.Get (n6)
        	vn6 = n.xyz
			App.feMeasureDistanceToPlane (vn6, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n6+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

        	n.Get (n7)
        	vn7 = n.xyz
			App.feMeasureDistanceToPlane (vn7, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n7+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

			n.Get (n8)
        	vn8 = n.xyz
			App.feMeasureDistanceToPlane (vn8, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n8+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 hexes.")
                GoTo Skip
            End If

			'Scenerio #1
            vec14(0) = vn4(0) - vn1(0)
            vec14(1) = vn4(1) - vn1(1)
            vec14(2) = vn4(2) - vn1(2)

			vec23(0) = vn2(0) - vn3(0)
            vec23(1) = vn2(1) - vn3(1)
            vec23(2) = vn2(2) - vn3(2)

			vec58(0) = vn8(0) - vn5(0)
            vec58(1) = vn8(1) - vn5(1)
            vec58(2) = vn8(2) - vn5(2)

			vec67(0) = vn6(0) - vn7(0)
            vec67(1) = vn6(1) - vn7(1)
            vec67(2) = vn6(2) - vn7(2)

            Int14 = App.feCoordOntoPlane (vn1, vec14, PlaneO, PlaneN, vint14)

            App.feMeasureDistance (vn4, vn1, edist)
            App.feMeasureDistance (vint14, vn1, edist1)
            App.feMeasureDistance (vint14, vn4, edist2)

            If edist1 < edist Then
                If edist2 < edist Then

				Int23 =  App.feCoordOntoPlane (vn3, vec23, PlaneO, PlaneN, vint23)

                App.feMeasureDistance (vn2, vn3, edist)
            	App.feMeasureDistance (vint23, vn2, edist1)
            	App.feMeasureDistance (vint23, vn3, edist2)

                If edist1 < edist Then
                	If edist2 < edist Then

					Int58 = App.feCoordOntoPlane (vn5, vec58, PlaneO, PlaneN, vint58)

                    App.feMeasureDistance (vn8, vn5, edist)
            		App.feMeasureDistance (vint58, vn5, edist1)
            		App.feMeasureDistance (vint58, vn8, edist2)

                	If edist1 < edist Then
                		If edist2 < edist Then

                    	Int67 = App.feCoordOntoPlane (vn7, vec67, PlaneO, PlaneN, vint67)

                        App.feMeasureDistance (vn6, vn7, edist)
            			App.feMeasureDistance (vint67, vn6, edist1)
            			App.feMeasureDistance (vint67, vn7, edist2)

                		If edist1 < edist Then
                			If edist2 < edist Then

                    		n.xyz = vint14
                    		nIntID14 = n.NextEmptyID
            				n.Put (nIntID14)
                            nset.Add (nIntID14)

            				n.xyz = vint23
            				nIntID23 = n.NextEmptyID
            				n.Put (nIntID23)
                            nset.Add (nIntID23)

                    		n.xyz = vint58
                    		nIntID58 = n.NextEmptyID
            				n.Put (nIntID58)
                            nset.Add (nIntID58)

            				n.xyz = vint67
            				nIntID67 = n.NextEmptyID
            				n.Put (nIntID67)
                            nset.Add (nIntID67)

                        	nsetp.Add (nIntID14)
                    		nsetp.Add (nIntID23)
                    		nsetp.Add (nIntID58)
                    		nsetp.Add (nIntID67)

                    		e.Node (0) = n1
                    		e.Node (1) = n2
                    		e.Node (2) = nIntID23
                    		e.Node (3) = nIntID14
                    		e.Node (4) = n5
                    		e.Node (5) = n6
                    		e.Node (6) = nIntID67
                    		e.Node (7) = nIntID58

                    		e.Put (eID)

							e.Node (0) = nIntID14
                    		e.Node (1) = nIntID23
                    		e.Node (2) = n3
                    		e.Node (3) = n4
                    		e.Node (4) = nIntID58
                    		e.Node (5) = nIntID67
                    		e.Node (6) = n7
                    		e.Node (7) = n8

                    		e.Put (e.NextEmptyID)

                    		nsetp.Clear

                            End If
                        End If
                        End If
                    End If
                    End If
                End If
                End If
            End If

            'Scenerio #2
            vec12(0) = vn2(0) - vn1(0)
            vec12(1) = vn2(1) - vn1(1)
            vec12(2) = vn2(2) - vn1(2)

			vec34(0) = vn3(0) - vn4(0)
            vec34(1) = vn3(1) - vn4(1)
            vec34(2) = vn3(2) - vn4(2)

			vec56(0) = vn6(0) - vn5(0)
            vec56(1) = vn6(1) - vn5(1)
            vec56(2) = vn6(2) - vn5(2)

			vec78(0) = vn7(0) - vn8(0)
            vec78(1) = vn7(1) - vn8(1)
            vec78(2) = vn7(2) - vn8(2)

			Int12 = App.feCoordOntoPlane (vn1, vec12, PlaneO, PlaneN, vint12)

            App.feMeasureDistance (vn1, vn2, edist)
            App.feMeasureDistance (vint12, vn1, edist1)
            App.feMeasureDistance (vint12, vn2, edist2)

            If edist1 < edist Then
                If edist2 < edist Then

				Int34 =  App.feCoordOntoPlane (vn3, vec34, PlaneO, PlaneN, vint34)

                App.feMeasureDistance (vn3, vn4, edist)
            	App.feMeasureDistance (vint34, vn3, edist1)
            	App.feMeasureDistance (vint34, vn4, edist2)

            	If edist1 < edist Then
                	If edist2 < edist Then

					Int56 = App.feCoordOntoPlane (vn5, vec56, PlaneO, PlaneN, vint56)

                    App.feMeasureDistance (vn5, vn6, edist)
            		App.feMeasureDistance (vint56, vn5, edist1)
            		App.feMeasureDistance (vint56, vn6, edist2)

            		If edist1 < edist Then
                		If edist2 < edist Then

                    	Int78 = App.feCoordOntoPlane (vn7, vec78, PlaneO, PlaneN, vint78)

                        App.feMeasureDistance (vn7, vn8, edist)
            			App.feMeasureDistance (vint78, vn7, edist1)
            			App.feMeasureDistance (vint78, vn8, edist2)

            			If edist1 < edist Then
                			If edist2 < edist Then

                    		n.xyz = vint12
                    		nIntID12 = n.NextEmptyID
            				n.Put (nIntID12)
                            nset.Add (nIntID12)

            				n.xyz = vint34
            				nIntID34 = n.NextEmptyID
            				n.Put (nIntID34)
                            nset.Add (nIntID34)

                    		n.xyz = vint56
                    		nIntID56 = n.NextEmptyID
            				n.Put (nIntID56)
                            nset.Add (nIntID56)

            				n.xyz = vint78
            				nIntID78 = n.NextEmptyID
            				n.Put (nIntID78)
                            nset.Add (nIntID78)

                        	nsetp.Add (nIntID12)
                    		nsetp.Add (nIntID34)
                    		nsetp.Add (nIntID56)
                    		nsetp.Add (nIntID78)

                    		e.Node (0) = n1
                    		e.Node (1) = nIntID12
                    		e.Node (2) = nIntID34
                    		e.Node (3) = n4
                        	e.Node (4) = n5
                    		e.Node (5) = nIntID56
                    		e.Node (6) = nIntID78
                    		e.Node (7) = n8

                    		e.Put (eID)

							e.Node (0) = nIntID12
                    		e.Node (1) = n2
                    		e.Node (2) = n3
                    		e.Node (3) = nIntID34
                        	e.Node (4) = nIntID56
                    		e.Node (5) = n6
                    		e.Node (6) = n7
                    		e.Node (7) = nIntID78

                    		e.Put (e.NextEmptyID)

                    		nsetp.Clear
							End If
                        End If
                        End If
                    End If
                    End If
                End If
            	End If
        	End If

			'Scenerio #3
            vec15(0) = vn1(0) - vn5(0)
            vec15(1) = vn1(1) - vn5(1)
            vec15(2) = vn1(2) - vn5(2)

			vec26(0) = vn2(0) - vn6(0)
            vec26(1) = vn2(1) - vn6(1)
            vec26(2) = vn2(2) - vn6(2)

			vec37(0) = vn3(0) - vn7(0)
            vec37(1) = vn3(1) - vn7(1)
            vec37(2) = vn3(2) - vn7(2)

			vec48(0) = vn4(0) - vn8(0)
            vec48(1) = vn4(1) - vn8(1)
            vec48(2) = vn4(2) - vn8(2)

			Int15 = App.feCoordOntoPlane (vn1, vec15, PlaneO, PlaneN, vint15)

            App.feMeasureDistance (vn1, vn5, edist)
            App.feMeasureDistance (vint15, vn1, edist1)
            App.feMeasureDistance (vint15, vn5, edist2)

            If edist1 < edist Then
                If edist2 < edist Then

				Int26 =  App.feCoordOntoPlane (vn2, vec26, PlaneO, PlaneN, vint26)

                App.feMeasureDistance (vn2, vn6, edist)
            	App.feMeasureDistance (vint26, vn2, edist1)
            	App.feMeasureDistance (vint26, vn6, edist2)

           		If edist1 < edist Then
                	If edist2 < edist Then

					Int37 = App.feCoordOntoPlane (vn3, vec37, PlaneO, PlaneN, vint37)

                    App.feMeasureDistance (vn3, vn7, edist)
            		App.feMeasureDistance (vint37, vn3, edist1)
            		App.feMeasureDistance (vint37, vn7, edist2)

            		If edist1 < edist Then
                		If edist2 < edist Then

                    	Int48 = App.feCoordOntoPlane (vn4, vec48, PlaneO, PlaneN, vint48)

                        App.feMeasureDistance (vn4, vn8, edist)
            			App.feMeasureDistance (vint48, vn4, edist1)
            			App.feMeasureDistance (vint48, vn8, edist2)

            			If edist1 < edist Then
                			If edist2 < edist Then

                    		n.xyz = vint15
                    		nIntID15 = n.NextEmptyID
            				n.Put (nIntID15)
                            nset.Add (nIntID15)

            				n.xyz = vint26
            				nIntID26 = n.NextEmptyID
            				n.Put (nIntID26)
                            nset.Add (nIntID26)

                    		n.xyz = vint37
                    		nIntID37 = n.NextEmptyID
            				n.Put (nIntID37)
                            nset.Add (nIntID37)

            				n.xyz = vint48
            				nIntID48 = n.NextEmptyID
            				n.Put (nIntID48)
                            nset.Add (nIntID48)

                        	nsetp.Add (nIntID15)
                    		nsetp.Add (nIntID26)
                    		nsetp.Add (nIntID37)
                    		nsetp.Add (nIntID48)

                    		e.Node (0) = n1
                    		e.Node (1) = n2
                    		e.Node (2) = n3
                    		e.Node (3) = n4
                        	e.Node (4) = nIntID15
                    		e.Node (5) = nIntID26
                    		e.Node (6) = nIntID37
                    		e.Node (7) = nIntID48

                    		e.Put (eID)

							e.Node (0) = nIntID15
                    		e.Node (1) = nIntID26
                    		e.Node (2) = nIntID37
                    		e.Node (3) = nIntID48
                        	e.Node (4) = n5
                    		e.Node (5) = n6
                    		e.Node (6) = n7
                    		e.Node (7) = n8

                    		e.Put (e.NextEmptyID)

                    		nsetp.Clear
							End If
                        End If
                        End If
                    End If
                    End If
                End If
            	End If
        	End If
        End If

        If e.topology = FTO_QUAD4 Then

			n1 = e.Node (0)
        	n2 = e.Node (1)
        	n3 = e.Node (2)
        	n4 = e.Node (3)

            nset.Add(n1)
            nset.Add(n2)
            nset.Add(n3)
            nset.Add(n4)

        	n.Get (n1)
        	vn1 = n.xyz
            App.feMeasureDistanceToPlane (vn1, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n1+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 quads.")
                GoTo Skip
            End If

        	n.Get (n2)
        	vn2 = n.xyz
            App.feMeasureDistanceToPlane (vn2, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n2+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 quads.")
                GoTo Skip
            End If

        	n.Get (n3)
        	vn3 = n.xyz
            App.feMeasureDistanceToPlane (vn3, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n3+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 quads.")
                GoTo Skip
            End If

			n.Get (n4)
        	vn4 = n.xyz
			App.feMeasureDistanceToPlane (vn4, PlaneO, PlaneN, dist)

            If Abs(dist) < 0.001 Then
            	App.feAppMessage (FCM_WARNING, "Node "+n4+ " on Element "+ eID + " is to close to slice plane.  Element will not be split into 2 quads.")
                GoTo Skip
            End If

			'Scenerio #1
            vec14(0) = vn4(0) - vn1(0)
            vec14(1) = vn4(1) - vn1(1)
            vec14(2) = vn4(2) - vn1(2)

			vec23(0) = vn2(0) - vn3(0)
            vec23(1) = vn2(1) - vn3(1)
            vec23(2) = vn2(2) - vn3(2)

            Int14 = App.feCoordOntoPlane (vn1, vec14, PlaneO, PlaneN, vint14)

            App.feMeasureDistance (vn4, vn1, edist)
            App.feMeasureDistance (vint14, vn1, edist1)
            App.feMeasureDistance (vint14, vn4, edist2)

            If edist1 < edist Then
                If edist2 < edist Then

					Int23 =  App.feCoordOntoPlane (vn3, vec23, PlaneO, PlaneN, vint23)

                	App.feMeasureDistance (vn2, vn3, edist)
            		App.feMeasureDistance (vint23, vn2, edist1)
            		App.feMeasureDistance (vint23, vn3, edist2)

                	If edist1 < edist Then
                		If edist2 < edist Then

                    		n.xyz = vint14
                    		nIntID14 = n.NextEmptyID
            				n.Put (nIntID14)
                            nset.Add(nIntID14)

            				n.xyz = vint23
            				nIntID23 = n.NextEmptyID
            				n.Put (nIntID23)
                            nset.Add(nIntID23)

                       		nsetp.Add (nIntID14)
                    		nsetp.Add (nIntID23)

                    		e.Node (0) = n1
                    		e.Node (1) = n2
                    		e.Node (2) = nIntID23
                    		e.Node (3) = nIntID14

                    		e.Put (eID)

							e.Node (0) = nIntID14
                    		e.Node (1) = nIntID23
                    		e.Node (2) = n3
                    		e.Node (3) = n4

                    		e.Put (e.NextEmptyID)

                    		nsetp.Clear

                   		End If
                	End If
            	End If
        	End If

            'Scenerio #2
            vec12(0) = vn2(0) - vn1(0)
            vec12(1) = vn2(1) - vn1(1)
            vec12(2) = vn2(2) - vn1(2)

			vec34(0) = vn3(0) - vn4(0)
            vec34(1) = vn3(1) - vn4(1)
            vec34(2) = vn3(2) - vn4(2)

			Int12 = App.feCoordOntoPlane (vn1, vec12, PlaneO, PlaneN, vint12)

            App.feMeasureDistance (vn1, vn2, edist)
            App.feMeasureDistance (vint12, vn1, edist1)
            App.feMeasureDistance (vint12, vn2, edist2)

            If edist1 < edist Then
                If edist2 < edist Then

				Int34 =  App.feCoordOntoPlane (vn3, vec34, PlaneO, PlaneN, vint34)

                App.feMeasureDistance (vn3, vn4, edist)
            	App.feMeasureDistance (vint34, vn3, edist1)
            	App.feMeasureDistance (vint34, vn4, edist2)

            		If edist1 < edist Then
                		If edist2 < edist Then

                    		n.xyz = vint12
                    		nIntID12 = n.NextEmptyID
            				n.Put (nIntID12)
                            nset.Add(nIntID12)

            				n.xyz = vint34
            				nIntID34 = n.NextEmptyID
            				n.Put (nIntID34)
                            nset.Add(nIntID34)

                        	nsetp.Add (nIntID12)
                    		nsetp.Add (nIntID34)

                    		e.Node (0) = n1
                    		e.Node (1) = nIntID12
                    		e.Node (2) = nIntID34
                    		e.Node (3) = n4

                    		e.Put (eID)

							e.Node (0) = nIntID12
                    		e.Node (1) = n2
                    		e.Node (2) = n3
                    		e.Node (3) = nIntID34

                    		e.Put (e.NextEmptyID)

                    		nsetp.Clear

                    	End If
                	End If
            	End If
        	End If
        End If

        Skip:

    	eID = eset.Next

    Wend

    Dim coincount As Long
    Dim coinnodes As Variant

    'App.feCheckCoincidentNode (nset.ID, nset.ID,1e-6, True, False, coincount,coinnodes)

    aView = App.Info_ActiveID (FT_VIEW)

    App.feWindowRegenerate (aView)

    nset.Clear

    App.feAppUnlock

    Done:

    App.feAppMessage (FCM_WARNING, "The 'Split mesh by plane' API has finished")

End Sub
