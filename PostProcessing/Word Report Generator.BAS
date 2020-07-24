'#Reference {00020905-0000-0000-C000-000000000046}#8.3#0#C:\Program Files\Microsoft Office\OFFICE11\MSWORD.OLB#Microsoft Word 11.0 Object Library#Word
Rem File: WordReportGenMultipleSolids.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim viewID As Long

	Dim solSet As femap.Set
	Set solSet = app.feSet

	Dim elemSet As femap.Set
	Set elemSet = app.feSet

	Dim elSolSet As femap.Set
	Set elSolSet = app.feSet

	Dim elPlateSet As femap.Set
	Set elPlateSet = app.feSet

	Dim nodeSet As femap.Set
	Set nodeSet = app.feSet

	Dim feProp As femap.Prop
	Set feProp = app.feProp

	Dim feMatl As femap.Matl
	Set feMatl = app.feMatl

	Dim feNode As femap.Node
	Set feNode = app.feNode

	Begin Dialog UserDialog 590,364,"User Information" ' %GRID:10,7,1,1
		text 40,35,110,21,"Name",.Text1
		text 40,70,110,21,"Company",.Text12
		text 40,105,110,21,"Organization",.Text13
		text 40,140,140,21,"Description:",.Text14
		TextBox 150,35,400,21,.TextBox1
		TextBox 150,70,400,21,.TextBox12
		TextBox 150,105,400,21,.TextBox13
		TextBox 40,161,510,161,.TextBox14,1
		OKButton 130,329,120,21
		CancelButton 340,329,120,21
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
	GoTo Bailout
	End If

	Dim feView As femap.View
	Set feView = app.feView
	Dim feView2 As femap.View
	Set feView2 = app.feView

	app.feAppGetActiveView( viewID )

	feView.Get(viewID)

	app.feAppGetActiveView(viewID)

	feView2.Get(viewID)

		For i = 0 To 21
		feView.Draw (i) = False
		Next i

		feView.Draw(26) = False
		feView.Draw (27) = False

		For i = 66 To 74
		feView.Draw (i) = False
		Next i

		For i = 79 To 85
		feView.Draw (i) = False
		Next i

		feView.Draw (91) = False

		feView.Draw (9) = True
		feView.Label (46) = 2
		feView.Label (0) = 4
		feView.Contour = FCF_CONTOUR
		feView.Deformed = FDF_DEFORMED
		feView.Draw (FVI_UNDEFORMED) = False
		feView.WindowBackColor = FCL_WHITE
		feView.WindowShadeBack = False
		feView.ColorMode (FVI_CONTOUR_LEGEND) = 1
		feView.color (FVI_CONTOUR_LEGEND) = FCL_BLACK
		feView.color(0) = FCL_BLACK
		feView.color(31) = FCL_BLACK
		feView.color(40) = FCL_BLACK

		rc = app.feFileProgramRun (False, True, False, "#SILENT(){~2006}<A-E><@18000><~0><@10501>10<OK>")

	'For Plates
	'feView.ContourData = 7033
	'For Solids
	feView.ContourData = 60031

	rc = feView.Put( viewID )
	rc = app.feViewRegenerate( 0 )

	Dim App2 As Word.Application
	Dim wrdDoc As Word.Document

    Set App2 = CreateObject("Word.Application")

    'create a new document

    Set wrdDoc = App2.Documents.Add

    'Dim crMsg As String

	crMsg  = vbCr

	Dim tNotes As String

	'Solid Properties
	Dim Area As Double
	Dim Volume As Double
	Dim cg As Variant
	Dim Inertia As Variant

	'View and Output Data

	Dim outSet As femap.OutputSet
	Set outSet = app.feOutputSet

	Dim dData As femap.output
	Set dData = app.feOutput

	Dim sData As femap.output
	Set sData = app.feOutput

	Dim sNData As femap.output
	Set sNData = app.feOutput

	Dim feElem As femap.Elem
	Set feElem = app.feElem

	'Elemental to Nodal Conversion Variables
	Dim approach As Long
	Dim from_SetID As Long
	Dim from_vectorID As Long
	Dim to_setID As Long
	Dim to_vectorID As Long
	Dim newVectorID As Long

	Dim feSolid As femap.Solid
	Set feSolid = app.feSolid

	Dim surfSet As femap.Set
	Set surfSet = app.feSet

	Dim nSolids As Long

	rc = app.feFilePictureCopy()

	Dim myRange As Word.Range
	Set myRange = App2.ActiveDocument.Paragraphs(1).Range

	Msg = "FEMAP Analysis Report"

	With myRange
	.Font.Bold = True
	.Font.size = 14
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .InsertAfter( crMsg )
    '.InsertAfter( crMsg )
    'Msg = "Model File Information"
    '.InsertAfter( Msg )
	Msg = vbCr
	'.InsertAfter( Msg )
	.InsertParagraphAfter
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'Write Name, Company, Organization, and Description to report

	Msg = dlg.TextBox1

	If Len (Msg) = 0 Then
	GoTo Company
	End If

	With myRange
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    '.InsertAfter( crMsg )
    Msg = "Name: "
    .InsertBefore( Msg )
    Msg = vbCr
    '.InsertAfter( Msg )
    .InsertParagraphAfter
	.Font.Bold = False
	.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Company:

	Msg = dlg.TextBox12

	If Len (Msg) = 0 Then
	GoTo Organization
	End If

	With myRange
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    '.InsertAfter( crMsg )
    Msg = "Company: "
    .InsertBefore( Msg )
    Msg = vbCr
    '.InsertAfter( Msg )
    .InsertParagraphAfter
    .Font.Bold = False
	.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Organization:

	Msg = dlg.TextBox13

	If Len (Msg) = 0 Then
	GoTo Description
	End If

	With myRange
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    '.InsertAfter( crMsg )
    Msg = "Organization: "
    .InsertBefore( Msg )
    Msg = vbCr
    .InsertAfter( Msg )
    .InsertParagraphAfter
    .Font.Bold = False
	.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Description:

	Msg = dlg.TextBox14

	If Len (Msg) = 0 Then
	GoTo Info
	End If

	With myRange
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    '.InsertAfter( crMsg )
    Msg = "Description: "
    .InsertBefore( Msg )
    Msg = vbCr
    .InsertAfter( Msg )
    .InsertParagraphAfter
    .Font.Bold = False
	.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Info:
	Msg = "Model File Information"

	With myRange
	.Font.Bold = True
	.Font.size = 14
	.InsertBefore Msg
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    .InsertAfter( crMsg )
    '.InsertAfter( crMsg )
    '.InsertAfter( Msg )
	'Msg = vbCr
	'.InsertAfter( Msg )
	.InsertParagraphAfter
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	' Get the Model File Name
	If Len( app.ModelName ) > 0 Then
		Msg = "This FEMAP Model is stored at: " + app.ModelName
	Else
		Msg = "This FEMAP Model has not been saved yet."
	End If
	With myRange
		.InsertBefore( Msg )
		.ParagraphFormat.Alignment = wdAlignParagraphLeft
    	.InsertAfter( crMsg )
		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'information on where the geometry came from
	With myRange
		Msg = "Geometry Import Information" + vbCr
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		rc = app.feNotesGet( 1, tNotes)
		.InsertAfter( tNotes )
		.InsertAfter( crMsg )
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	rc = solSet.Clear()
	rc = solSet.AddAll( FT_SOLID )

	With myRange
		Msg = "Model contains " + Format$( solSet.count, "0" ) + " solids."
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	While feSolid.Next()

		rc = feSolid.MassProp( Area, Volume, cg, Inertia )
		rc = surfSet.Clear()

		'find out how many faces on the solid
		rc = surfSet.AddRule( feSolid.ID, FGD_SURFACE_ONSOLID )

		rc = elemSet.Clear
		rc = elemSet.AddRule( feSolid.ID, FGD_ELEM_ATSOLID )

		If elemSet.count Then
			rc = feElem.Get( elemSet.First() )
			rc = feProp.Get( feElem.propID )
			rc = feMatl.Get( feProp.matlID )
			matMsg = "Solid meshed with " + Format$( elemSet.count, "0" ) + " elements connected to Material " + Format$(feProp.matlID, "0") + " Titled - " + feMatl.title
		Else
			matMsg = ""
		End If


		'display the mass properties
		With myRange
			Msg = "Solid ID " + Format$( feSolid.ID, "0") + " consists of " + Format$( surfSet.count, "0") + " faces. "
			.InsertAfter( Msg )
			Msg = "Solid Total Volume = " + Format$( Volume, ".00000" ) + vbCr
			.InsertAfter( Msg )
			Msg = "Solid C.G. x = " + Format$( cg(0), ".00000" ) + " y = " + Format$( cg(1), ".00000" ) + " z = " + Format$(cg(2), ".00000") + vbCr
			.InsertAfter( Msg )
			.InsertAfter( matMsg )
			.InsertAfter( crMsg )
			.InsertParagraphAfter
			.Font.Bold = False
			.Font.size = 12
		End With

		'reset the range to the end of the document
		myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Wend

	'Paste the contour picture in
	With myRange
		.Paste
		.InsertAfter ( crMsg )
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'create the picture of deformation

	app.feAppGetActiveView( viewID )

	feView.Get(viewID)

	conlevelmode = feView.Label (FVI_CONTOUR_LEVELS)

	feView.Label (FVI_CONTOUR_LEVELS) = 0

	feView.ContourData = 1

	rc = feView.Put( viewID )
	rc = app.feViewRegenerate( 0 )

	Wait (1)

	'Copy and Paste the deformation picture in

	rc = app.feFilePictureCopy()

	'back to contour

	feView.Label (FVI_CONTOUR_LEVELS) = conlevelmode

	'For Plates
	'feView.ContourData = 7033
	'For Solids
	feView.ContourData = 60031

	rc = feView.Put( viewID )
	rc = app.feViewRegenerate( 0 )

	With myRange
		.Paste
		.InsertAfter ( crMsg )
		.Font.Bold = False
		.Font.size = 12
	End With

	' Mesh Information
	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	' get node and element count
	tet10count = 0
	other_elemCount = 0

	rc = feElem.Reset

	While feElem.Next
		If feElem.topology = FTO_TETRA10 Then
			tet10count = tet10count + 1
		Else
			other_elemCount = other_elemCount + 1
		End If
	Wend

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'display the mesh properties
	With myRange
		.InsertAfter (crMsg)
		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		Msg = "Mesh Information"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'display the mesh properties
	With myRange
		Msg = "Mesh consists of " + Format$( tet10count, "0") + " 10-noded Tetra Elements"
		.InsertAfter( Msg )
		Msg = " and " + Format$( other_elemCount, "0" ) + " other elements." + vbCr
		.InsertAfter( Msg )
		Msg = "Total Node Count = " + Format$( app.Info_Count( FT_NODE), "0" )  + vbCr
		.InsertAfter( Msg )
		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With


'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		Msg = "Material Information"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	rc = feMatl.Reset

	While feMatl.Next
		'display material info

		With myRange
			Msg = "Material ID " + Format$( feMatl.ID, "0" ) + " titled - " + feMatl.title
			.InsertAfter( Msg )
			.InsertAfter( crMsg )
			Msg = "Modulus of Elasticity Specified as " + Format$( feMatl.Ex,"00.0E+00") + vbCr
			.InsertAfter( Msg )
			Msg = "Shear Modulus Specified as " + Format$( feMatl.Gx,"00.0E+00") + vbCr
			.InsertAfter( Msg )
			Msg = "Poisson's Ration Specified as " + Format$( feMatl.Nuxy,"0.000") + vbCr
			.InsertAfter( Msg )
			Msg = "Limit Tension Stress Specified as " + Format$( feMatl.TensionLimit1 ,"00.0E+00") + vbCr
			.InsertAfter( Msg )
			Msg = "Limit Compression Stress Specified as " + Format$( feMatl.CompressionLimit1  ,"00.0E+00") + vbCr
			.InsertAfter( Msg )
			Msg = "Limit Shear Stress Specified as " + Format$( feMatl.ShearLimit  ,"00.0E+00") + vbCr
			.InsertAfter( Msg )
			.InsertParagraphAfter
			.Font.Bold = False
			.Font.size = 12
		End With

		'reset the range to the end of the document
		myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	Wend

	' Some applied Loads Information

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		Msg = "Applied Loads"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'display applied load info

	Dim feLoadGeom As femap.LoadGeom
	Set feLoadGeom = app.feLoadGeom

	With myRange
		While feLoadGeom.Next
			If feLoadGeom.setID = app.Info_ActiveID( FT_LOAD_DIR ) Then
				Select Case feLoadGeom.type
					Case FLT_SEPRESSURE
						Msg = "A Pressure Load of " + Format$( feLoadGeom.LoadValue, "0.000") + " is applied to Surface" + Str$( feLoadGeom.geomID )
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
					Case FLT_SNFORCE
						If feLoadGeom.CSys = 0 Then
							If feLoadGeom.dirMode = FLD_NONE Then
								If Abs( feLoadGeom.load(0) ) > 0 Then
									Msg = "A  Force of " + Format$(feLoadGeom.load(0), "0.000") + " in the Global X Direction is applied to Surface" + Str$( feLoadGeom.geomID )
									.InsertAfter( Msg )
								End If
								If Abs( feLoadGeom.load(1) ) > 0 Then
									Msg = "A  Force of " + Format$(feLoadGeom.load(1), "0.000") + " in the Global Y Direction is applied to Surface" + Str$( feLoadGeom.geomID )
									.InsertAfter( Msg )
								End If
								If Abs( feLoadGeom.load(2)) > 0 Then
									Msg = "A  Force of " + Format$(feLoadGeom.load(2), "0.000") + " in the Global Z Direction is applied to Surface" + Str$( feLoadGeom.geomID )
									.InsertAfter( Msg )
								End If
								.InsertAfter( crMsg )
							End If
							If feLoadGeom.dirMode = FLD_NORMALTOSURFACE Then
								If Abs( feLoadGeom.load(0) ) > 0 Then
									Msg = "A  Force of " + Format$(feLoadGeom.load(0), "0.000") + " is applied normal to Surface" + Str$( feLoadGeom.geomID )
									.InsertAfter( Msg )
								End If
								.InsertAfter( crMsg )
							End If
						End If
				End Select
			End If

		Wend

		'Add in Nodal Loads
		Dim feLoadMesh As femap.LoadMesh
		Set feLoadMesh = app.feLoadMesh

		While feLoadMesh.Next
			If feLoadMesh.setID = app.Info_ActiveID( FT_LOAD_DIR ) Then
				If feLoadMesh.type  = FLT_NFORCE  Then
					Msg = "A Nodal Force is applied in CSys " + Str$(feLoadMesh.CSys) + " of X = " +Format$( feLoadMesh.load(0), "0.000") + " of Y = " +Format$( feLoadMesh.load(1), "0.000") +" of Z = " +Format$( feLoadMesh.load(2), "0.000")
					.InsertAfter( Msg )
				End If
			End If
		Wend

		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		.InsertAfter ( crMsg )
		Msg = "Constraints"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	'display applied load info

	Dim feBCGeom As femap.BCGeom
	Set feBCGeom = app.feBCGeom

	With myRange
		While feBCGeom.Next
			If feBCGeom.setID = app.Info_ActiveID( FT_BC_DIR ) Then
				'Surface based constraints
				If feBCGeom.geomType = FT_SURFACE Then
					Select Case feBCGeom.type
					Case 0
						'Pinned or Fixed
						'For solid geometry, there are no rotational DOFs, therefore Pinned and Fixed are the same
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained from translating in any direction"
							.InsertAfter( Msg)
							.InsertAfter( crMsg )
					Case 1 'Arbitrary in coordinate system
						If feBCGeom.outCSys = -1 Then
							Csysmes = "the nodes' Output Coordinate Systems"
						Else
							Csysmes = "Coordinate System" + Str$(feBCGeom.outCSys )
						End If
						'TX Only
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RX degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RY Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RY degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RZ Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RZ degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and TY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX and TY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and TZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX and TZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and TZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY and TZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RX and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RX and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RY and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the RY and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
					Case 2
						'Normal To Surface (Symmetry)
						Msg = "All the nodes of Surface" +Str$(feBCGeom.geomID) + " are constrained normal To Surface (symmetry) "
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
					Case 5
						'Cylindrical
						Msg = "All the nodes of Surface" + Str$(feBCGeom.geomID) + " have cylindrical constraints, "
						If feBCGeom.dof(0) Then
							Msg = Msg + "constraining radial growth, "
						Else
							Msg = Msg + "allowing radial growth, "
						End If
						If feBCGeom.dof(1) Then
							Msg = Msg + "constraining  rotation around the central axis, "
						Else
							Msg = Msg + "allowing rotation around the central axis, "
						End If
						If feBCGeom.dof(2) Then
							Msg = Msg + "and constraining sliding along the central axis."
						Else
							Msg = Msg + "and allowing sliding along the central axis."
						End If
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
					End Select
				End If
				'Curve Based constraints
				If feBCGeom.geomType = FT_CURVE Then
					Select Case feBCGeom.type
					Case 0		'Pinned, Fixed, or No Rotation
						'Fixed Constraint
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are fixed (constrained in all degrees of freedom)"
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'Pinned Constraint
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are pinned (constrained in all translational degrees of freedom)"
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'No Rotation Constraint
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in all rotational degrees of freedom"
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
					Case 1 'Arbitrary in coordinate system
						If feBCGeom.outCSys = -1 Then
							Csysmes = "the nodes' Output Coordinate Systems"
						Else
							Csysmes = "Coordinate System" + Str$(feBCGeom.outCSys )
						End If
						'TX Only
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RX degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RY Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RY degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RZ Only
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RZ degree of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and TY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX and TY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and TZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX and TZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and TZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY and TZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RX and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RX and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RX and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'RY and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the RY and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, RY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TZ, RY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RX
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, TY, TZ, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, TY, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, and RY
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TX, RX, RY, and RZ
						If feBCGeom.dof(0) = True And feBCGeom.dof(1) = False And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TX, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RX
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RX degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, TZ, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TY, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = True And feBCGeom.dof(2) = False And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TY, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, and RY
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = False Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, and RY degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = False And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = False And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
						'TZ, RX, RY, and RZ
						If feBCGeom.dof(0) = False And feBCGeom.dof(1) = False And feBCGeom.dof(2) = True And feBCGeom.dof(3) = True And feBCGeom.dof(4) = True And feBCGeom.dof(5) = True Then
							Msg = "All the nodes of Curve" + Str$(feBCGeom.geomID ) + " are constrained in the TZ, RX, RY, and RZ degrees of freedom in " + Csysmes
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
						End If
					End Select
				End If
			End If
		Wend

		.InsertParagraphAfter
		.Font.Bold = False
		.Font.size = 12
	End With

		'display applied BC on node info

	Dim feBCNode As femap.BCNode
	Set feBCNode = app.feBCNode

	With myRange
		While feBCNode.Next
			If feBCNode.setID = app.Info_ActiveID( FT_BC_DIR ) Then
				'Fixed Constraint
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " is fixed (all degrees of freedom constrained)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'Pinned Constraint
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " is pinned (all translational degrees of freedom constrained)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'No Rotation Constraint
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has all rotational degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX Only
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY Only
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ Only
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RX Only
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RX degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RY Only
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RY degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RZ Only
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RZ degree of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX and TY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX and TY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX and TZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX and TZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX and RX
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY and TZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY and TZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY and RX
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ and RX
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RX and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RX and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RX and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RX and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'RY and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has RY and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, and RX
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, and RZ degrees of freedom constrained (Z Anti-symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, RX, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, RX, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, RX, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, RX, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, RX, RY, and Rz degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ and RX
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, and RY degrees of freedom constrained (Y Anti-symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ, RX, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, RX, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ, RX, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TZ, RX, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TZ, RX, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, and RX
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, RX, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RX degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, RX, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, TY, TZ, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, TY, TZ, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, RX, and RY
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, RX, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, RX, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, RY, and RZ degrees of freedom constrained (X Symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TX, RX, RY, and RZ
				If feBCNode.dof(0) = True And feBCNode.dof(1) = False And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TX, RX, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ and RX
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, and RX degrees of freedom constrained (X Anti-symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ, RX, and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, RX, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ, RX, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, TZ, RX, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, TZ, RX, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, RX, and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, RX, and RY degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, RX, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, RX, and RZ degrees of freedom constrained (Y Symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TY, RX, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = True And feBCNode.dof(2) = False And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TY, RX, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ, RX, and RY
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = False Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ, RX, and RY degrees of freedom constrained (Z Symmetry)"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ, RX, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = False And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = False And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ, RY, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
				'TZ, RX, RY, and RZ
				If feBCNode.dof(0) = False And feBCNode.dof(1) = False And feBCNode.dof(2) = True And feBCNode.dof(3) = True And feBCNode.dof(4) = True And feBCNode.dof(5) = True Then
						Msg = "Node" + Str$(feBCNode.ID ) + " has TZ, RY, RX, and RZ degrees of freedom constrained"
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
				End If
			End If
		Wend

	.InsertParagraphAfter
	.Font.Bold = False
	.Font.size = 12
	End With

	'now some detailed stress information
	rc  = app.feAppGetActiveView( viewID )
	rc = feView.Get( viewID )

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
	With myRange
		Msg = "Analysis Results Information all Elements/Materials"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	rc = outSet.Get( feView.OutputSet )
	dData.setID = feView.OutputSet
	dData.Get( feView.DeformData )
	sData.setID = feView.OutputSet
	sData.Get(feView.ContourData)

	sNData.ID = 8999999
	sNData.ID = sNData.NextEmptyID


	' Want to find the maximum nodal stress, must convert the data vector from elemental to nodal

	approach = 	0 'Average
	from_SetID = app.Info_ActiveID(FT_OUT_CASE )
	from_vectorID = feView.ContourData
	to_setID = from_SetID

	to_vectorID = sNData.ID

    rc = app.feOutputProcessConvert( approach, from_SetID, from_vectorID, from_SetID, to_vectorID, 0 )

	sNData.setID = feView.OutputSet
	sNData.Get( sNData.ID)

	Dim ndID As Long
	Dim maxStress As Double
	Dim maxndID As Long
	Dim FS As Double

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	With myRange
		Msg = "Output Set Title: " + outSet.title
		.InsertAfter ( Msg )
		.InsertAfter ( crMsg )
		Msg = "Maximum Displacement of " + Format$( dData.absmaxval, ".00000" ) + " at Node " + Format$( dData.maxvalID, "0" )
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		Msg = "Maximum Elemental Centroidal Stress of " + Format$( sData.maxval , ".00000" ) + " at Elem " + Format$( sData.maxvalID, "0" )
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		Msg = "Maximum Nodal Stress of " + Format$( sNData.maxval, ".00000") + " at Node " + Format$( sNData.maxvalID, "0" )
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		Msg  = ""
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.Font.Bold = False
		.Font.size = 12
	End With

	' now for each solid
	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
	With myRange
		Msg = "Analysis Results For Each Solids"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With

	rc = feSolid.Reset

	While feSolid.Next
		rc = nodeSet.Clear
		rc = elemSet.Clear

		rc = elemSet.AddRule( feSolid.ID, FGD_ELEM_ATSOLID )
		rc = nodeSet.AddSetRule( elemSet.ID, FGD_NODE_ONELEM )

		If nodeSet.count Then
			rc = feElem.Get( elemSet.First() )
			rc = feProp.Get( feElem.propID )
			rc = feMatl.Get( feProp.matlID )

			maxStress = 0.0
			maxndID = 0

			rc = nodeSet.Reset
			ndID = nodeSet.First
			While ndID
				If sNData.value(ndID) > maxStress Then
					maxStress = sNData.value(ndID)
					maxndID = ndID
				End If
				ndID = nodeSet.Next
			Wend

			If maxndID Then
				rc = feNode.Get( maxndID )
				'reset the range to the end of the document
				myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
				With myRange
						Msg = "For Solid " + Format$( feSolid.ID, "0" ) + " Material " + Format$( feMatl.ID, "0" )
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg = "Maximum Nodal Stress of " + Format$( maxStress,  ".00000" )+ " is at Node " + Format$( maxndID, "0" ) +"."
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg = "This Maximum vonMises Stress was found at global X = " + Format$( feNode.x, "0.0000" ) + ", Y = " + Format$( feNode.y, "0.0000") + ", Z = " + Format$( feNode.z, "0.0000")
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg  = ""
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
						.Font.Bold = False
						.Font.size = 12
				End With

				If feMatl.TensionLimit1 = 0 Then
					With myRange
						Msg = "No Material Limit Tension Stress specified, impossible to calculate Factor of Safety"
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						.Font.Bold = False
						.Font.size = 12
					End With
				Else
					If Abs(maxStress) = 0.0 Then
						With myRange
							Msg = "No Nodal Maximum Stress available, impossible to calculate Factor of Safety"
							.InsertAfter ( crMsg )
							.Font.Bold = False
							.Font.size = 12
						End With
					Else
						FS =feMatl.TensionLimit1 / maxStress
						With myRange
							Msg = "For Solid " + Format$( feSolid.ID, "0" ) + " Material " + Format$( feMatl.ID, "0" ) + " with a Yield Tensile Strength of " + Format$( feMatl.TensionLimit1, "0.0000" )
							.InsertAfter ( Msg )
							.InsertAfter ( crMsg )
							Msg = "The Factor of Safety is " + Format$( FS, "0.0000")
							.InsertAfter ( Msg )
							.InsertAfter ( crMsg )
							Msg  = ""
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
							.Font.Bold = False
							.Font.size = 12
						End With
						If FS < 1.0 Then
							myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
							With myRange
								Msg = "!!! WARNING, THIS FACTOR OF SAFETY INDICATES THE MATERIAL HAS YIELDED !!!"
								.InsertAfter( Msg )
								.InsertAfter( crMsg )
								.InsertAfter( crMsg )
								.Font.Bold = True
								.Font.size = 12
								End With
						End If
					End If
				End If
			End If
		End If
	Wend

	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
	With myRange
	    .InsertAfter ( crMsg )
		Msg = "Analysis Results For Each Material"
		.InsertAfter( Msg )
		.InsertAfter( crMsg )
		.InsertAfter( crMsg )
		.Font.Bold = True
		.Font.size = 14
	End With


	rc = feMatl.Reset
	While feMatl.Next

		rc = nodeSet.Clear
		rc = elemSet.Clear

		rc = elemSet.AddRule( feMatl.ID, FGD_ELEM_BYMATL )
		rc = nodeSet.AddSetRule( elemSet.ID, FGD_NODE_ONELEM )

		If nodeSet.count Then
			rc = feElem.Get( elemSet.First() )
			rc = feProp.Get( feElem.propID )
			rc = feMatl.Get( feProp.matlID )

			maxStress = 0.0
			maxndID = 0

			rc = nodeSet.Reset
			ndID = nodeSet.First
			While ndID
				If sNData.value(ndID) > maxStress Then
					maxStress = sNData.value(ndID)
					maxndID = ndID
				End If
				ndID = nodeSet.Next
			Wend

			If maxndID Then
				rc = feNode.Get( maxndID )
				'reset the range to the end of the document
				myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
				With myRange
						Msg = "For Material " + Format$( feMatl.ID, "0" ) + " Material " + Format$( feMatl.ID, "0" )
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg = "Maximum Nodal Stress of " + Format$( maxStress,  ".00000" )+ " is at Node " + Format$( maxndID, "0" ) +"."
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg = "This Maximum vonMises Stress was found at global X = " + Format$( feNode.x, "0.0000" ) + ", Y = " + Format$( feNode.y, "0.0000") + ", Z = " + Format$( feNode.z, "0.0000")
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						Msg  = ""
						.InsertAfter( Msg )
						.InsertAfter( crMsg )
						.Font.Bold = False
						.Font.size = 12
				End With

				If feMatl.TensionLimit1 = 0 Then
					With myRange
						Msg = "No Material Limit Tension Stress specified, impossible to calculate Factor of Safety"
						.InsertAfter ( Msg )
						.InsertAfter ( crMsg )
						.Font.Bold = False
						.Font.size = 12
					End With
				Else
					If Abs(maxStress) = 0.0 Then
						With myRange
							Msg = "No Nodal Maximum Stress available, impossible to calculate Factor of Safety"
							.InsertAfter ( crMsg )
							.Font.Bold = False
							.Font.size = 12
						End With
					Else
						FS =feMatl.TensionLimit1 / maxStress
						With myRange
							Msg = "For Material " + Format$( feMatl.ID, "0" ) + " Material " + Format$( feMatl.ID, "0" ) + " with a Yield Tensile Strength of " + Format$( feMatl.TensionLimit1, "0.0000" )
							.InsertAfter ( Msg )
							.InsertAfter ( crMsg )
							Msg = "The Factor of Safety is " + Format$( FS, "0.0000")
							.InsertAfter ( Msg )
							.InsertAfter ( crMsg )
							Msg  = ""
							.InsertAfter( Msg )
							.InsertAfter( crMsg )
							.Font.Bold = False
							.Font.size = 12
						End With
						If FS < 1.0 Then
							myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End
							With myRange
								Msg = "!!! WARNING, THIS FACTOR OF SAFETY INDICATES THE MATERIAL HAS YIELDED !!!"
								.InsertAfter( Msg )
								.InsertAfter( crMsg )
								.InsertAfter( crMsg )
								.Font.Bold = True
								.Font.size = 12
								End With
						End If
					End If
				End If
			End If
		End If
	Wend


'Now delete the temporary Data Vector

	rc = app.feDeleteOutput( app.Info_ActiveID( FT_OUT_CASE ), -sNData.ID )

	'reset the range to the end of the document
	myRange.SetRange Start:=App2.ActiveDocument.Content.End, End:=App2.ActiveDocument.Content.End

	feView.Get (viewID)

	rc = feView2.Put( viewID )

	rc = app.feViewRegenerate( 0 )

	App2.visible = True

	BailOut:

End Sub
