Rem File: LayupThicknessToOutput.BAS
Sub Main
	Dim outSetID As Long
	'Dim elemType As Integer

'---------------------------------------------------------------------------
	Const maxNoElems As Long = 100000       ' max number of elements
	elemType = FET_L_PLATE          	' Element Type to process
	outVectorID = 400000			'Output  Vector ID
' --------------------------------------------------------------------------

	Dim App As femap.model
    Set App = feFemap()

	Dim ol As femap.Set
	Set ol = App.feSet
	Dim feProp As femap.Prop

    Dim Lay As femap.Layup
    Set Lay = App.feLayup

	Dim rc As Long
	Dim el As femap.Elem
	Set el = App.feElem()

	Dim Eid As Long

	' Output set vars
	Dim output0 As femap.Output
	Set output0 = App.feOutput

	Dim setID As Long
	Dim outSet As femap.OutputSet
	Set outSet = App.feOutputSet

	Dim elIDV As Variant
	Dim thicknessV As Variant

	Dim status As Long
	Dim count As Long
	Dim elemCount As Long

	Dim elLIST(maxNoElems) As Long
	Dim thick(maxNoElems) As Double

	Set feProp = App.feProp

	rc = ol.Select(FT_ELEM, True, "Choose Elements for Thickness Criteria")
	If rc = -1 Then
		count = ol.Count()
		If (count > maxNoElems) Then
			Dim msgs As String
			msgs = "Error, You must Select less than "+CStr(maxNoElems)+ " Elements at a time."
	    		MsgBox (msgs)

		ElseIf (count < maxNoElems And count <> 0) Then

			' count used by status bar,  num elems + output set tasks
			count = count + 3
			status = 1

			j = App.feAppStatusShow(True, count)
			App.feAppStatusUpdate (status)
			j = App.feAppStatusRedraw()

			Eid = ol.Next()
			i = 0

			Do While Eid > 0
				rc = el.Get(Eid)
				If el.type = FET_L_LAMINATE_PLATE Then
					rc = feProp.Get( el.propID )
					Lay.Get (feProp.layupID)
                    thick(i) = Lay.LayupThickness
					'thick(i) = feProp.pval(0)

				   	status = status + 1
	   			   	App.feAppStatusUpdate (status)
	    			elLIST(i) = Eid
			    	i = i + 1
   				 End If
			    Eid = ol.Next()
			Loop

			elemCount = i
			rc = App.feAppUnlock()
			j = App.feAppStatusRedraw()

			' load elem list and thicknesses into Variants
			elIDV = elLIST
			thicknessV = thick

			'Create the output set
			setID = outSet.NextEmptyID()
   			outSet.title = "Layup Thickness Values"
   			outSet.value = 0
   			outSet.analysis = 0
   			outSet.Put (setID)

			'Create the vector in the empty set
			rc = output0.InitScalarAtElem(setID, outVectorID, "Layup Thickness", 4, True)

		    	status = status + 1
			App.feAppStatusUpdate (status)
   			j = App.feAppStatusRedraw()

			rc = output0.PutScalarAtElem(elemCount, elIDV, thicknessV)
			rc = output0.Put(-1)

			status = status + 1
   			App.feAppStatusUpdate (status)
   			j = App.feAppStatusRedraw()

			' Issue Message
   			Dim Str As String
   			Dim sval As String
   			Dim color As Long
   			color = 2

			Str = "Output Set "

			sval = Str + CStr(setID)
   			Str = " created for Layup Thickness data"
   			sval = sval + Str

			rc = App.feAppMessage(color, sval)

		End If
	End If

	j = App.feAppStatusShow(False, 4)
	j = App.feAppStatusRedraw()

End Sub




	
