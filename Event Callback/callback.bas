Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim V As femap.Var
	Set V = App.feVar

	Dim pr As femap.Prop
	Set pr = App.feProp

	Dim aset As femap.AnalysisMgr
	Set aset = App.feAnalysisMgr

	Dim pID As Long
	Dim Ptype As Long

	Dim K As Double
	Dim K1 As Double
	Dim K2 As Double
	Dim Iter As Double

	Dim i As Long
	Dim Kval As Long

	i = V.GetVarID ("ITER")
	rc = V.Get (i)
	Iter = V.value
	i = V.GetVarID ("K1")
	rc = V.Get (i)
	K1 = V.value
	i = V.GetVarID ("K2")
	rc = V.Get (i)
	K2 = V.value
	i = V.GetVarID ("PID")
	rc = V.Get (i)
	pID = V.value
	i = V.GetVarID ("K")
	Kval = V.Get (i)
	K = V.value

	pr.Get (pID)

	Ptype = pr.type

	K = K + Iter

	If K>=K1 And K<=K2 Then

	pr.type = Ptype

	pr.pval(0) = K

	pr.Put(pID)

	V.value = K

	V.Put (i)

	aset.Get (1)
	aset.title()  = "Callback Val=" + Str( K )
	aset.Put( 1 )

	aset.Analyze (1)

	Else

	App.feAppEventCallback (FEVENT_RESULTSEND, "")

	End If

	
End Sub
