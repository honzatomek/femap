Sub Main
	Dim App As femap.model
    Set App = feFemap()

Dim s As femap.Set
Dim v As femap.Set
Dim vc As femap.Set
Set vc = App.feSet
Dim e As femap.Set
Set e = App.feSet
Dim ov As femap.Output
Set ov = App.feOutput
Dim ow As femap.Output
Set ow = App.feOutput
Dim minID As Long
Dim maxID As Long
Dim minVAL As Double
Dim maxVAL As Double
Dim numVec As Long
Dim x() As Double
Dim id() As Long
Dim total_SET() As Long
Dim total_ID() As Long
Dim total_title() As String
Dim nc As Long
Dim maxtitle As String
Dim rc As Long
Dim bMax As Boolean
Dim total_pre As String


Dim t As femap.DataTable
Set t = App.feDataTable()

App.feAppManagePanes ("Data Table", 1)
t.Lock( False )
t.Clear()

If	App.feSelectOutput( "Select Output Vectors",0,  FOT_ANY, FOC_ANY, FT_ELEM, False, s, v ) = FE_OK  Then
	If e.Select( FT_ELEM, True, "Select Elements" ) = FE_OK Then

		rc = App.feAppMessageBox( 2, "Ok to Build Table of Maximum Values (No=Minimum) ?" )
		If rc = FE_OK Then
		   bMax = True
		Else
		   bMax = False
		End If

		vc.Copy(v.ID)
		numVec = v.Count()
		ReDim x(numVec)
		ReDim id(numVec)
		ReDim total_SET(numVec)
		ReDim total_ID(numVec)
		ReDim total_title(numVec)
		For i=0 To numVec-1
			id(i) = i+1
		Next i

	    ' Start looking at vectors................................
		i = 0
		While v.Next()
			total_minSET = 0
			total_minID = 0
			total_minVAL = 1.0E30
			total_maxSET = 0
			total_maxID = 0
			total_maxVAL = -1.0E30
			' Look across all selected Sets........................
			s.Reset()
			While s.Next()
				ov.GetFromSet( s.CurrentID, v.CurrentID )
				ov.FindMaxMin( e.ID, False, minID, minVAL, maxID, maxVAL )
				If ( minVAL < total_minVAL ) Then
						total_minVAL = minVAL
						total_minID = minID
						total_minSET = s.CurrentID
				End If
				If ( maxVAL > total_maxVAL ) Then
						total_maxVAL = maxVAL
						total_maxID = maxID
						total_maxSET = s.CurrentID
				End If
			Wend

			'Totals now contain min/max values across all sets.................
			'Get all the results for total set and min/max entity ID.............
			If bMax Then
				total_SET(i) = total_maxSET
				total_ID(i) = total_maxID
				total_title(i) = ov.title
				total_pre = "Max "
			Else
		    	total_SET(i) = total_minSET
				total_ID(i) = total_minID
				total_title(i) = ov.title
				total_pre = "Min "
			End If

			vc.Reset()
			j = 0
			While vc.Next()
				ow.GetFromSet( total_SET(i), vc.CurrentID )
				x(j) = ow.Value(total_ID(i))
				j = j + 1
			Wend
			maxtitle = total_pre + ov.title
			t.AddColumn( False, False, ov.location, 0, maxtitle, FCT_DOUBLE, numVec, id, x, nc )
			i = i+1
		Wend

		t.AddColumn( False, False, ov.location, 0, "Title", FCT_STRING, numVec, id, total_title, nc )
		t.SetColumnPosition( nc, t.FindColumn("ID"),  True  )
		t.AddColumn( False, False, ov.location, 0, "Output Set", FCT_INT, numVec, id, total_SET, nc )
		t.SetColumnPosition( nc, t.FindColumn("Title"),  True  )

		For i=0 To numVec-1
			id(i) = i
		Next i
		t.UpdateColumn( t.FindColumn("ID"), FCT_INT, numVec, id, total_ID )

		For i=0 To numVec-1
			id(0) = i
			t.SetBackgroundColor( 3+i, 1, id, 255, 255, 192 )
		Next i
		t.ClearSelection()

	End If
End If
End Sub
