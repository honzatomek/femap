Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim pr As femap.Prop
	Set pr = App.feProp

	Dim allPropS As femap.Set	'All Properties in the Model
	Set allPropS = App.feSet
	allPropS.Reset

	Dim pset As femap.Set	'Plate Properties in the Model
	Set pset = App.feSet

	Dim pID As Long
	pID = allPropS.First
	While pID <> 0
		pr.Get ( pID )
		If pr.type < 17 Or pr.type > 18	Then	'Property is a Plate
			pset.Add ( pID )
		End If
	Wend

	rc = pset.SelectID ( FT_PROP, "Select Shell Property to Vary", pID )
	If rc = 2 Or rc = 4 Then
		End
	End If

	Dim Var As femap.Var
	Set Var = App.feVar

	Dim Ptype As Long

	Dim K As Double
	Dim K1 As Double
	Dim K2 As Double
	Dim Num As Integer
	Dim Iter As Double
	Dim Range As Double

	Begin Dialog UserDialog 280,161,"Shell Thickness Variation" ' %GRID:10,7,1,1
		text 30,21,80,14,"Start Thickness",.Text1
		text 30,56,80,14,"End Thickness",.Text12
		text 30,91,80,14,"Iterations",.Text13
		TextBox 130,14,110,21,.valstart
		TextBox 130,49,110,21,.valend
		TextBox 130,84,110,21,.iter
		OKButton 30,119,80,28
		CancelButton 130,119,110,28
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
		End
	End If

	K1 = Val (dlg.valstart)

	K2 = Val (dlg.valend)

	Num = Val (dlg.iter)

	Range = K2 - K1

	Iter = Range/(Num-1)

	K = K1

	rc = Var.Define ("PID", Str(pID))
	rc = Var.Define ("K1", Str(K1))
	rc = Var.Define ("K2", Str(K2))
	rc = Var.Define ("ITER", Str(Iter))
	rc = Var.Define ("K", Str(K))

    Dim path As String
    path = MacroDir + "\callback.bas"

	' Set up Femap
	App.Pref_OutputSetTitles = 1
	
	rc = App.feAppEventCallback (FEVENT_RESULTSEND,path )

	Dim aset As femap.AnalysisMgr
	Set aset = App.feAnalysisMgr

	pr.Get (pID)

	Ptype = pr.type

	If K >=K1 And K <= K2 Then

	pr.type = Ptype

	pr.pval(0) = K

	pr.Put(pID)

	aset.Get (1)
	aset.title()  = "Callback Val=" + Str( K )
	aset.Put( 1 )

	aset.Analyze (1)

	Else

	App.feAppEventCallback (FEVENT_RESULTSEND, "")

	End If

End Sub
