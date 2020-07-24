Rem File: HoneyComb_PShell_Property.bas
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim m As femap.Matl
	Set m = App.feMatl

	Dim nMatl As Long
	Dim vID As Variant
	Dim vTitle As Variant
	Dim FaceID As Long
	Dim CoreID As Long

	If App.Info_Count(FT_MATL) = 0 Then
		Msg = "No Materials in Model, Exiting..."
		rc = MsgBox( Msg, vbOkOnly )
		GoTo OK
	End If

	m.GetTitleList (0,0,nMatl,vID,vTitle)

	For I=0 To nMatl-1

   vTitle(I) = Str$(vID(I)) + ".." + vTitle(I)

   Next I

	Dim FaceMat$()
	FaceMat$() = vTitle

	Dim CoreMat$()
	CoreMat$() = vTitle

	Begin Dialog UserDialog 600,448,"Honeycomb Panel Cross Section (PSHELL) Property Input" ' %GRID:10,7,1,1
		OKButton 50,378,150,49
		CancelButton 400,378,150,49
		text 50,49,260,35,"Total Face Sheet Thickness (T = t/2 * 2 in figure)",.Text12
		text 50,126,260,21,"Core Thickness (D in figure)",.Text13
		TextBox 330,49,200,21,.FaceThick
		TextBox 330,196,200,21,.CoreDens
		TextBox 330,126,200,21,.CoreThick
		Picture 30,238,540,126,MacroDir$+"\Xsection.bmp",0,.picture1
		text 50,84,260,21,"Face Sheet Material",.Text1
		text 50,196,260,21,"Core Density (Rho)",.Text14
		text 50,161,260,21,"Core Material",.Text15
		text 50,21,260,14,"Property Tiltle",.Text16
		TextBox 330,14,200,21,.PropTitle
		DropListBox 330,84,200,21,FaceMat(),.FaceMat,1
		DropListBox 330,161,200,21,CoreMat(),.CoreMat,1

	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
	GoTo OK
	End If

	Dim FaceThickWork As Double
    FaceThickWork = Val (dlg.FaceThick)
    Dim CoreThickWork As Double
    CoreThickWork = Val (dlg.CoreThick)
	Dim CoreDensWork As Double
    CoreDensWork = Val (dlg.CoreDens)

    If FaceThickWork <= 0 Then
    GoTo FAIL
    End If

    If CoreThickWork <= 0 Then
    GoTo FAIL
    End If

    If CoreDensWork <= 0 Then
    GoTo FAIL
    End If

    FaceID = Val( Split(dlg.FaceMat, "..",2) (0) )
    CoreID = Val( Split(dlg.CoreMat, "..",2) (0) )


Dim propID As Long

Dim feProp As femap.Prop

Set feProp = App.feProp

propID = feProp.NextEmptyID

feProp.type = 17

feProp.matlID = FaceID
feProp.pval (0) = FaceThickWork
feProp.pval (7) = CoreDensWork*CoreThickWork
feProp.pval (8) = (CoreThickWork+FaceThickWork/2)/2
feProp.pval (9) = -(CoreThickWork+FaceThickWork/2)/2
feProp.pval (10) = (12 * ((FaceThickWork*(CoreThickWork^2))/4))/FaceThickWork^3
feProp.pval (11) = CoreThickWork/FaceThickWork
feProp.ExtraMatlID (0) = FaceID
feProp.ExtraMatlID (1) = CoreID
feProp.title = dlg.PropTitle

feProp.Put(propID)

If propID > feProp.Last Then

GoTo GOOD

End If

FAIL:

rc = App.feAppMessage (3,"Errors have occured, no property created")

GoTo OK

GOOD:

rc = App.feAppMessage (0, "Honeycomb Property Created")

OK:

End Sub

