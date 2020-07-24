'www.PredictiveEngineering.com
'All Rights Reserved, 2014
'Predictive Engineering Assumes No Responsibility For Results Obtained From API
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'API Written by Predictive Engineering, Rev-19 (Tested on Femap v11.1.0)
'Rev-19 (updated by Adrian Jensen, January 2014) removes unnecessary output vectors and the option to combine output sets via the API.
'It is now much faster to combine output sets with Femap and then use the API on the resulting combination.
'The program formatting has also been updated and cleaned up.

' This API started out as a simple mid-plane stress calculator and the original program was provided by a colleague of the Femap development team around 1998.
' Since that date, the program has been expanded to calculate the ASME Stress Intensity values For plate, beam And 8-Node brick elements.

' For selected output case this program calculates and creates new output vectors:
' - Plate Membrane X Normal Stress (300000)
' - Plate Membrane Y Normal Stress (300005)
' - Plate Membrane XY Shear Stress (300010)
' - Plate Membrane MajorPrn Stress (300015)
' - Plate Membrane MinorPrn Stress (300020)
' - Plate Membrane Stress Intensity (300025)
' - Plate Membrane VonMises Stress (300030)

' - Solid Stress Intensity (300050)
' - Solid Triaxial Stress (300055)

' - Beam EndA Axial Stress (302400)
' - Beam EndB Axial Stress (302500)

' This program also calculates and overwrites existing output vectors:
' - Plate Top Stress Intensity (7031)
' - Plate Bot Stress Intensity (7431)
' - Plate Top Triaxial Stress (7030)
' - Plate Bot Triaxial Stress (7430)

' Reference ASME Section VIII, Div. 2, Appendix 4.

' Please feel free to suggest any improvements to www.PredictiveEngineering.com

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Sub Main
Dim App As femap.model
Set App = feFemap()

Dim OutputSet As femap.OutputSet
Set OutputSet = App.feOutputSet

Dim feoutput As femap.Output
Set feoutput = App.feOutput

Dim OutputSetID As Long
Dim OutputSetIDlast As Long
Dim nElem As Long
Dim maxCorner As Long
Dim c(8) As Double
Dim c1 As Variant
Dim c2 As Variant
Dim c3 As Variant
Dim c4 As Variant
Dim cenVal1 As Variant
Dim cenVal11(2) As Double
Dim j As Long								' variable for nElem - looping through all elements
Dim i As Long

Dim minIDos As Long
Dim maxIDos As Long
Dim numos As Long

Dim os As femap.Set
Set os = App.feSet
os.Select(28,True,"Select Output Sets for Data Processing")

minIDos = os.First
maxIDos = os.Last
numos = os.Count

rc = App.feAppMessage( 0, "Number of Selected Output Sets =" + Str(numos))

'Status Bar
Dim plateelems As Long
plateelems = 0
Dim solidelems As Long
solidelems = 0
Dim beamelems As Long
beamelems = 0
Dim Element As femap.Elem
Set Element = App.feElem
While Element.Next
	If Element.type = FET_L_PLATE Then
		plateelems = plateelems + 1
	ElseIf Element.type = FET_L_SOLID Then
		solidelems = solidelems + 1
	ElseIf Element.type = FET_L_BEAM Then
		beamelems = beamelems + 1
	End If
Wend

Dim Scount As Long
Scount = numos*(plateelems*11 + solidelems*2 + beamelems*2)
Dim status As Long
status = 0
Dim bigstatus As Long
bigstatus = 1000
App.feAppStatusShow(True, Scount)
App.feAppStatusUpdate (status)
App.feAppStatusRedraw()


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                                                    P L A T E   M E M B R A N E    S T R E S S E S
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' (7020+7420)/2-->300000
' (Plate Top X Normal Stress + Plate Bot X Normal Stress)/2 --> Plate Membrane X Normal Stress

OutputSetID = os.First

While OutputSetID > 0

	rc = feoutput.InitElemWithCorner(OutputSetID, 7020, 100220, 150220, 200220, 250220, 0, 0, 0, 0,"Plate Top X Normal Stress",0,True)
	If feoutput.Get(7020) Then		' Plate Top X Normal Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, ptxc, ptxc1, ptxc2, ptxc3, ptxc4, ptxc5, ptxc6, ptxc7, ptxc8)
		rc = feoutput.InitElemWithCorner(OutputSetID, 7420, 100620, 150620, 200620, 250620, 0, 0, 0, 0,"Plate Bot X Normal Stress",0,True)
		rc = feoutput.Get(7420)		' Plate Bot X Normal Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, pbxc, pbxc1, pbxc2, pbxc3, pbxc4, pbxc5, pbxc6, pbxc7, pbxc8)
		rc = feoutput.Get(300000)

		ReDim xmmsc(nElem) As Double
		ReDim xmmsc1(nElem) As Double
		ReDim xmmsc2(nElem) As Double
		ReDim xmmsc3(nElem) As Double
		ReDim xmmsc4(nElem) As Double

		For j = 0 To nElem-1
			xmmsc(j) = (ptxc(j) + pbxc(j))/2
			xmmsc1(j) = (ptxc1(j)+pbxc1(j))/2
			xmmsc2(j) = (ptxc2(j)+pbxc2(j))/2
			xmmsc3(j) = (ptxc3(j)+pbxc3(j))/2
			xmmsc4(j) = (ptxc4(j)+pbxc4(j))/2

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 100, 101, 102, 103, 104, 0, 0, 0, 0,"Plate Membrane X Normal Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, xmmsc, xmmsc1, xmmsc2, xmmsc3, xmmsc4, 0, 0, 0, 0 )
		rc = feoutput.Put(300000)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' (7021+7421)/2-->300005
' (Plate Top Y Normal Stress + Plate Bot Y Normal Stress)/2 --> Plate Membrane Y Normal Stress

		rc = feoutput.InitElemWithCorner(OutputSetID, 7021, 100221, 150221, 200221, 250221, 0, 0, 0, 0,"Plate Top Y Normal Stress",0,True)
		rc = feoutput.Get(7021)		' Plate Top Y Normal Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, ptyc, ptyc1, ptyc2, ptyc3, ptyc4, ptyc5, ptyc6, ptyc7, ptyc8)
		rc = feoutput.InitElemWithCorner(OutputSetID, 7421, 100621, 150621, 200621, 250621, 0, 0, 0, 0,"Plate Bot Y Normal Stress",0,True)
		rc = feoutput.Get(7421)		' Plate Bot Y Normal Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, pbyc, pbyc1, pbyc2, pbyc3, pbyc4, pbyc5, pbyc6, pbyc7, pbyc8)
		rc = feoutput.Get(300005)

		ReDim ymmsc(nElem) As Double
		ReDim ymmsc1(nElem) As Double
		ReDim ymmsc2(nElem) As Double
		ReDim ymmsc3(nElem) As Double
		ReDim ymmsc4(nElem) As Double

		For j = 0 To nElem-1
			ymmsc(j) = (ptyc(j) + pbyc(j))/2
			ymmsc1(j) = (ptyc1(j)+pbyc1(j))/2
			ymmsc2(j) = (ptyc2(j)+pbyc2(j))/2
			ymmsc3(j) = (ptyc3(j)+pbyc3(j))/2
			ymmsc4(j) = (ptyc4(j)+pbyc4(j))/2

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 105, 106, 107, 108, 109, 0, 0, 0, 0,"Plate Membrane Y Normal Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, ymmsc, ymmsc1, ymmsc2, ymmsc3, ymmsc4, ymmsc5, ymmsc6, ymmsc7, ymmsc8 )
		rc = feoutput.Put(300005)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' (7023+7423)/2-->300010
' (Plate Top XY Shear Stress + Plate Bot XY Shear Stress)/2 --> Plate Membrane XY Shear Stress

		rc = feoutput.InitElemWithCorner(OutputSetID, 7023, 100223, 150223, 200223, 250223, 0, 0, 0, 0,"Plate Top XY Shear Stress",0,True)
		rc = feoutput.Get(7023)		' Plate Top XY Shear Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, ptxyc, ptxyc1, ptxyc2, ptxyc3, ptxyc4, ptxyc5, ptxyc6, ptxyc7, ptxyc8)

		rc = feoutput.InitElemWithCorner(OutputSetID, 7423, 100623, 150623, 200623, 250623, 0, 0, 0, 0,"Plate Bot XY Shear Stress",0,True)
		rc = feoutput.Get(7423)		' Plate Bot XY Shear Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, pbxyc, pbxyc1, pbxyc2, pbxyc3, pbxyc4, pbxyc5, pbxyc6, pbxyc7, pbxyc8)
		rc = feoutput.Get(300010)

		ReDim xymmsc(nElem) As Double
		ReDim xymmsc1(nElem) As Double
		ReDim xymmsc2(nElem) As Double
		ReDim xymmsc3(nElem) As Double
		ReDim xymmsc4(nElem) As Double

		For j = 0 To nElem-1
			xymmsc(j) = (ptxyc(j) + pbxyc(j))/2
			xymmsc1(j) = (ptxyc1(j)+pbxyc1(j))/2
			xymmsc2(j) = (ptxyc2(j)+pbxyc2(j))/2
			xymmsc3(j) = (ptxyc3(j)+pbxyc3(j))/2
			xymmsc4(j) = (ptxyc4(j)+pbxyc4(j))/2

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 110, 111, 112, 113, 114, 0, 0, 0, 0,"Plate Membrane XY Shear Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, xymmsc, xymmsc1, xymmsc2, xymmsc3, xymmsc4, xymmsc5, xymmsc6, xymmsc7, xymmsc8 )
		rc = feoutput.Put(300010)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ((300000+300005)/2)+sqr(((300000-300005)^2)/4+300010^2)-->300015
' Plate Membrane X Normal Stress, Plate Membrane Y Normal Stress, Plate Membrane XY Shear --> Plate Membrane MajorPrn Stress

		rc = feoutput.Get(300015)

		ReDim mjpmmsc(nElem) As Double
		ReDim mjpmmsc1(nElem) As Double
		ReDim mjpmmsc2(nElem) As Double
		ReDim mjpmmsc3(nElem) As Double
		ReDim mjpmmsc4(nElem) As Double

		For j = 0 To nElem-1
			mjpmmsc(j) = ((xmmsc(j)+ymmsc(j))/2)+(Sqr((xmmsc(j)-ymmsc(j))^2/4 + xymmsc(j)^2))
			mjpmmsc1(j) = ((xmmsc1(j)+ymmsc1(j))/2)+(Sqr((xmmsc1(j)-ymmsc1(j))^2/4 + xymmsc1(j)^2))
			mjpmmsc2(j) = ((xmmsc2(j)+ymmsc2(j))/2)+(Sqr((xmmsc2(j)-ymmsc2(j))^2/4 + xymmsc2(j)^2))
			mjpmmsc3(j) = ((xmmsc3(j)+ymmsc3(j))/2)+(Sqr((xmmsc3(j)-ymmsc3(j))^2/4 + xymmsc3(j)^2))
			mjpmmsc4(j) = ((xmmsc4(j)+ymmsc4(j))/2)+(Sqr((xmmsc4(j)-ymmsc4(j))^2/4 + xymmsc4(j)^2))

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 115, 116, 117, 118, 119, 0, 0, 0, 0,"Plate Membrane MajorPrn Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, mjpmmsc, mjpmmsc1, mjpmmsc2, mjpmmsc3, mjpmmsc4, mjpmmsc5, mjpmmsc6, mjpmmsc7, mjpmmsc8 )
		rc = feoutput.Put(300015)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ((300000+300005)/2)-sqr(((300000-300005)^2)/4+300010^2)-->300020
' Plate Membrane X Normal Stress, Plate Membrane Y Normal Stress, Plate Membrane XY Shear --> Plate Membrane MinorPrn Stress

		rc = feoutput.Get(300020)

		ReDim mnpmmsc(nElem) As Double
		ReDim mnpmmsc1(nElem) As Double
		ReDim mnpmmsc2(nElem) As Double
		ReDim mnpmmsc3(nElem) As Double
		ReDim mnpmmsc4(nElem) As Double

		For j = 0 To nElem-1
			mnpmmsc(j) = ((xmmsc(j)+ymmsc(j))/2)-(Sqr((xmmsc(j)-ymmsc(j))^2/4 + xymmsc(j)^2))
			mnpmmsc1(j) = ((xmmsc1(j)+ymmsc1(j))/2)-(Sqr((xmmsc1(j)-ymmsc1(j))^2/4 + xymmsc1(j)^2))
			mnpmmsc2(j) = ((xmmsc2(j)+ymmsc2(j))/2)-(Sqr((xmmsc2(j)-ymmsc2(j))^2/4 + xymmsc2(j)^2))
			mnpmmsc3(j) = ((xmmsc3(j)+ymmsc3(j))/2)-(Sqr((xmmsc3(j)-ymmsc3(j))^2/4 + xymmsc3(j)^2))
			mnpmmsc4(j) = ((xmmsc4(j)+ymmsc4(j))/2)-(Sqr((xmmsc4(j)-ymmsc4(j))^2/4 + xymmsc4(j)^2))

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 120, 121, 122, 123, 124, 0, 0, 0, 0,"Plate Membrane MinorPrn Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, mnpmmsc, mnpmmsc1, mnpmmsc2, mnpmmsc3, mnpmmsc4, mnpmmsc5, mnpmmsc6, mnpmmsc7, mnpmmsc8 )
		rc = feoutput.Put(300020)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' max(|300015|, |300020|, |300015-300020|)-->300025
' Plate Membrane MajorPrn Stress, Plate Membrane MinorPrn Stress -->  Plate Membrane Stress Intensity

		rc = feoutput.Get(300025)

		ReDim c1(nElem) As Double
		ReDim c2(nElem) As Double
		ReDim c3(nElem) As Double
		ReDim c4(nElem) As Double
		ReDim c5(nElem) As Double
		ReDim c6(nElem) As Double
		ReDim c7(nElem) As Double
		ReDim c8(nElem) As Double
		ReDim cenVal1(nElem) As Double

		For j = 0 To nElem-1
			c(0) = mjpmmsc1(j)
			c(1) = mnpmmsc1(j)
			c(2) = mjpmmsc2(j)
			c(3) = mnpmmsc2(j)
			c(4) = mjpmmsc3(j)
			c(5) = mnpmmsc3(j)
			c(6) = mjpmmsc4(j)
			c(7) = mnpmmsc4(j)
			cenVal11(0) = mjpmmsc(j)
			cenVal11(1) = mnpmmsc(j)

			While i < 7
				A = Abs(c(i))					     '|Major Principle Stress|
				B = Abs(c(i+1))					  '|Minor Principle Stress|
				diff = Abs(c(i) - c(i+1))		'|Major Principle Stress - Minor Principle Stress|
				If diff > A Then
					If diff > B Then
						c(i) = diff
					Else
						c(i) = B
					End If
				ElseIf A > B Then
					c(i) = A
				ElseIf B > A Then
					c(i) = B
				End If
				i = i +2
			Wend
			i = 0


			A = Abs(cenVal11(0))
			B = Abs(cenVal11(1))
			diff = Abs(cenVal11(0) - cenVal11(1))
			If diff > A Then
				If diff > B Then
					cenVal11(0) = diff
				Else
					cenVal11(0) = B
				End If
			ElseIf A > B Then
				cenVal11(0) = A
			ElseIf B > A Then
				cenVal11(0) = B
			End If

			c1(j) = c(0)
			c2(j) = c(2)
			c3(j) = c(4)
			c4(j) = c(6)
			cenVal1(j) = cenVal11(0)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 125, 126, 127, 128, 129, 0, 0, 0, 0,"Plate Membrane Stress Intensity",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(300025)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' sqr(300015^2-300015*300020+300020^2)-->300030
' Plate Membrane MajorPrn Stress, Plate Membrane MinorPrn Stress -->  Plate Membrane VonMises Stress

		rc = feoutput.Get(300030)

		ReDim vmmmsc(nElem) As Double
		ReDim vmmmsc1(nElem) As Double
		ReDim vmmmsc2(nElem) As Double
		ReDim vmmmsc3(nElem) As Double
		ReDim vmmmsc4(nElem) As Double

		For j = 0 To nElem-1
			vmmmsc(j) = Sqr(mjpmmsc(j)^2 - mjpmmsc(j)*mnpmmsc(j) + mnpmmsc(j)^2)
			vmmmsc1(j) = Sqr(mjpmmsc1(j)^2 - mjpmmsc1(j)*mnpmmsc1(j) + mnpmmsc1(j)^2)
			vmmmsc2(j) = Sqr(mjpmmsc2(j)^2 - mjpmmsc2(j)*mnpmmsc2(j) + mnpmmsc2(j)^2)
			vmmmsc3(j) = Sqr(mjpmmsc3(j)^2 - mjpmmsc3(j)*mnpmmsc3(j) + mnpmmsc3(j)^2)
			vmmmsc4(j) = Sqr(mjpmmsc4(j)^2 - mjpmmsc4(j)*mnpmmsc4(j) + mnpmmsc4(j)^2)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 130, 131, 132, 133, 134, 0, 0, 0, 0,"Plate Membrane VonMises Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, vmmmsc, vmmmsc1, vmmmsc2, vmmmsc3, vmmmsc4, vmmmsc5, vmmmsc6, vmmmsc7, vmmmsc8 )
		rc = feoutput.Put(300030)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                                                    P L A T E    S U R F A C E    S T R E S S E S
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' max(|7026|, |7027|, |7026-7027|)-->7031
' Plate Top MajorPrn Stress, Plate Top MinorPrn Stress -->  Plate Top Stress Intensity

		rc = feoutput.InitElemWithCorner(OutputSetID, 7026, 100226, 150226, 200226, 250226, 0, 0, 0, 0,"Plate Top MajorPrn Stress",0,True)
		rc = feoutput.Get(7026)					' Plate Top MajorPrn Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, ptmajpc, ptmajpc1, ptmajpc2, ptmajpc3, ptmajpc4, ptmajpc5, ptmajpc6, ptmajpc7, ptmajpc8)
		rc = feoutput.InitElemWithCorner(OutputSetID, 7027, 100227, 150227, 200227, 250227, 0, 0, 0, 0,"Plate Top MinorPrn Stress",0,True)
		rc = feoutput.Get(7027) 				' Plate Top MinorPrn Stress
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, ptminpc, ptminpc1, ptminpc2, ptminpc3, ptminpc4, ptminpc5, ptminpc6, ptminpc7, ptminpc8)
		rc = feoutput.Get(7031)
		feoutput.title = "Plate Top Stress Intensity"

		ReDim c1(nElem) As Double
		ReDim c2(nElem) As Double
		ReDim c3(nElem) As Double
		ReDim c4(nElem) As Double
		ReDim c5(nElem) As Double
		ReDim c6(nElem) As Double
		ReDim c7(nElem) As Double
		ReDim c8(nElem) As Double
		ReDim cenVal1(nElem) As Double

		For j = 0 To nElem-1
			c(0) = ptmajpc1(j)
			c(1) = ptminpc1(j)
			c(2) = ptmajpc2(j)
			c(3) = ptminpc2(j)
			c(4) = ptmajpc3(j)
			c(5) = ptminpc3(j)
			c(6) = ptmajpc4(j)
			c(7) = ptminpc4(j)
			cenVal11(0) = ptmajpc(j)
			cenVal11(1) = ptminpc(j)

			While i < 7
				A = Abs(c(i))
				B = Abs(c(i+1))
				diff = Abs(c(i) - c(i+1))
				If diff > A Then
					If diff > B Then
						c(i) = diff
					Else
						c(i) = B
					End If
				ElseIf A > B Then
					c(i) = A
				ElseIf B > A Then
					c(i) = B
				End If
				i = i +2
			Wend
			i = 0

			A = Abs(cenVal11(0))
			B = Abs(cenVal11(1))
			diff = Abs(cenVal11(0) - cenVal11(1))
			If diff > A Then
				If diff > B Then
					cenVal11(0) = diff
				Else
					cenVal11(0) = B
				End If
			ElseIf A > B Then
				cenVal11(0) = A
			ElseIf B > A Then
				cenVal11(0) = B
			End If

			c1(j) = c(0)
			c2(j) = c(2)
			c3(j) = c(4)
			c4(j) = c(6)
			cenVal1(j) = cenVal11(0)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 135, 136, 137, 138, 139, 0, 0, 0, 0,"Plate Top Stress Intensity",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(7031)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' max(|7426|, |7427|, |7426-7427|)-->7431
' Plate Bot MajorPrn Stress, Plate Bot MinorPrn Stress -->  Plate Bot Stress Intensity

		rc = feoutput.InitElemWithCorner(OutputSetID, 7426, 100626, 150626, 200626, 250626, 0, 0, 0, 0,"Plate Bot MajorPrn",0,True)
		rc = feoutput.Get(7426)
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, pbmajpc, pbmajpc1, pbmajpc2, pbmajpc3, pbmajpc4, pbmajpc5, pbmajpc6, pbmajpc7, pbmajpc8)
		rc = feoutput.InitElemWithCorner(OutputSetID, 7427, 100627, 150627, 200627, 250627, 0, 0, 0, 0,"Plate Bot MinorPrn",0,True)
		rc = feoutput.Get(7427)
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, pbminpc, pbminpc1, pbminpc2, pbminpc3, pbminpc4, pbminpc5, pbminpc6, pbminpc7, pbminpc8)
		rc = feoutput.Get(7431)
		feoutput.title = "Plate Bottom Stress Intensity"

		For j = 0 To nElem-1
			c(0) = pbmajpc1(j)
			c(1) = pbminpc1(j)
			c(2) = pbmajpc2(j)
			c(3) = pbminpc2(j)
			c(4) = pbmajpc3(j)
			c(5) = pbminpc3(j)
			c(6) = pbmajpc4(j)
			c(7) = pbminpc4(j)
			cenVal11(0) = pbmajpc(j)
			cenVal11(1) = pbminpc(j)

			While i < 7
				A = Abs(c(i))
				B = Abs(c(i+1))
				diff = Abs(c(i) - c(i+1))
				If diff > A Then
					If diff > B Then
						c(i) = diff
					Else
						c(i) = B
					End If
				ElseIf A > B Then
					c(i) = A
				ElseIf B > A Then
					c(i) = B
				End If
				i = i +2
			Wend
			i = 0

			A = Abs(cenVal11(0))
			B = Abs(cenVal11(1))
			diff = Abs(cenVal11(0) - cenVal11(1))
			If diff > A Then
				If diff > B Then
					cenVal11(0) = diff
				Else
					cenVal11(0) = B
				End If
			ElseIf A > B Then
				cenVal11(0) = A
			ElseIf B > A Then
				cenVal11(0) = B
			End If
			c1(j) = c(0)
			c2(j) = c(2)
			c3(j) = c(4)
			c4(j) = c(6)
			cenVal1(j) = cenVal11(0)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 140, 141, 142, 143, 144, 0, 0, 0, 0,"Plate Bot Stress Intensity",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(7431)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 7026+7027-->7030
' Plate Top MajorPrn Stress, Plate Top MinorPrn Stress -->  Plate Top Triaxial Stress

		rc = feoutput.InitElemWithCorner(OutputSetID, 145, 146, 147, 148, 149, 0, 0, 0, 0,"Plate Top Triaxial Stress",0,True)
		rc = feoutput.Get(7030)
		feoutput.title = "Plate Top Triaxial Stress"

		For j = 0 To nElem-1
			c(0) = ptmajpc1(j)
			c(1) = ptminpc1(j)
			c(2) = ptmajpc2(j)
			c(3) = ptminpc2(j)
			c(4) = ptmajpc3(j)
			c(5) = ptminpc3(j)
			c(6) = ptmajpc4(j)
			c(7) = ptminpc4(j)
			cenVal11(0) = ptmajpc(j)
			cenVal11(1) = ptminpc(j)

			c1(j) = c(0) + c(1)
			c2(j) = c(2) + c(3)
			c3(j) = c(4) + c(5)
			c4(j) = c(6) + c(7)
			cenVal1(j) = cenVal11(0) + cenVal11(1)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 145, 146, 147, 148, 149, 0, 0, 0, 0,"Plate Top Triaxial Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(7030)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 7426+7427-->7430
' Plate Bot MajorPrn Stress, Plate Bot MinorPrn Stress -->  Plate Bot Triaxial Stress

		rc = feoutput.InitElemWithCorner(OutputSetID, 150, 151, 152, 153, 154, 0, 0, 0, 0,"Plate Bot Triaxial Stress",0,True)
		rc = feoutput.Get(7430)
		feoutput.title = "Plate Bot Triaxial Stress"

		For j = 0 To nElem-1
			c(0) = pbmajpc1(j)
			c(1) = pbminpc1(j)
			c(2) = pbmajpc2(j)
			c(3) = pbminpc2(j)
			c(4) = pbmajpc3(j)
			c(5) = pbminpc3(j)
			c(6) = pbmajpc4(j)
			c(7) = pbminpc4(j)
			cenVal11(0) = pbmajpc(j)
			cenVal11(1) = pbminpc(j)

			c1(j) = c(0) + c(1)
			c2(j) = c(2) + c(3)
			c3(j) = c(4) + c(5)
			c4(j) = c(6) + c(7)
			cenVal1(j) = cenVal11(0) + cenVal11(1)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 150, 151, 152, 153, 154, 0, 0, 0, 0,"Plate Bot Triaxial Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(7430)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		rc = App.feAppMessage( 0, "Output Vectors of Plates for Output case" + Str(OutputSetID) + " are done")
	Else
		rc = App.feAppMessage( 0, "There are no Output Vectors of Plates for Output case" + Str(OutputSetID))
	End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                                                                 S O L I D    S T R E S S E S
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' max(|60016|, |60017|, |60018|)-->3000050
' Solid Max Prin Stress, Solid Min Prin Stress, Solid Int Prin Stress -->  Solid Stress Intensity

	Dim holder(3) As Variant
	Dim chooser As Variant
	chooser = 0

	If feoutput.Get(60016) Then
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, smaxpc, smaxpc1, smaxpc2, smaxpc3, smaxpc4, smaxpc5, smaxpc6, smaxpc7, smaxpc8)
		rc = feoutput.Get(60017)
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, sminpc, sminpc1, sminpc2, sminpc3, sminpc4, sminpc5, sminpc6, sminpc7, sminpc8)
		rc = feoutput.Get(60018)
		rc = feoutput.GetElemWithCorner( nElem, maxCorner, eID, sintpc, sintpc1, sintpc2, sintpc3, sintpc4, sintpc5, sintpc6, sintpc7, sintpc8)
		rc = feoutput.Get(300050)

		ReDim s1(nElem) As Double
		ReDim s2(nElem) As Double
		ReDim s3(nElem) As Double
		ReDim s4(nElem) As Double
		ReDim s5(nElem) As Double
		ReDim s6(nElem) As Double
		ReDim s7(nElem) As Double
		ReDim s8(nElem) As Double
		ReDim senVal1(nElem) As Double

		For j = 0 To nElem-1

			holder(0) = Abs(smaxpc(j) - sminpc(j))			'This section will calculate the Solid Stress Intensity for the centroidal data
			holder(1) = Abs(smaxpc(j) - sintpc(j))
			holder(2) = Abs(sminpc(j) - sintpc(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			senVal1(j) = chooser

			holder(0) = Abs(smaxpc1(j) - sminpc1(j))			'This section will calculate the Solid Stress Intensity for the first node set
			holder(1) = Abs(smaxpc1(j) - sintpc1(j))
			holder(2) = Abs(sminpc1(j) - sintpc1(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s1(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc2(j) - sminpc2(j))			'This section will calculate the Solid Stress Intensity for the second node set
			holder(1) = Abs(smaxpc2(j) - sintpc2(j))
			holder(2) = Abs(sminpc2(j) - sintpc2(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s2(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc3(j) - sminpc3(j))			'This section will calculate the Solid Stress Intensity for the third node set
			holder(1) = Abs(smaxpc3(j) - sintpc3(j))
			holder(2) = Abs(sminpc3(j) - sintpc3(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s3(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc4(j) - sminpc4(j))			'This section will calculate the Solid Stress Intensity for the forth node set

			holder(1) = Abs(smaxpc4(j) - sintpc4(j))
			holder(2) = Abs(sminpc4(j) - sintpc4(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s4(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc5(j) - sminpc5(j))			'This section will calculate the Solid Stress Intensity for the fifth node set
			holder(1) = Abs(smaxpc5(j) - sintpc5(j))
			holder(2) = Abs(sminpc5(j) - sintpc5(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s5(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc6(j) - sminpc6(j))			'This section will calculate the Solid Stress Intensity for the sixth node set
			holder(1) = Abs(smaxpc6(j) - sintpc6(j))
			holder(2) = Abs(sminpc6(j) - sintpc6(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s6(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc7(j) - sminpc7(j))			'This section will calculate the Solid Stress Intensity for the seventh node set
			holder(1) = Abs(smaxpc7(j) - sintpc7(j))
			holder(2) = Abs(sminpc7(j) - sintpc7(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s7(j) = chooser
			chooser = 0

			holder(0) = Abs(smaxpc8(j) - sminpc8(j))			'This section will calculate the Solid Stress Intensity for the eighth node set
			holder(1) = Abs(smaxpc8(j) - sintpc8(j))
			holder(2) = Abs(sminpc8(j) - sintpc8(j))
			For i = 0 To 2
				If	holder(i) > chooser Then
					chooser = holder(i)
				End If
			Next i
			s8(j) = chooser
			chooser = 0

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 300050, 161, 162, 163, 164, 165, 166, 167, 168,"Solid Stress Intensity",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, senVal1, s1, s2, s3, s4, s5, s6, s7, s8 )
		rc = feoutput.Put(300050)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 60016+60017+60018-->3000055
' Solid Max Prin Stress, Solid Min Prin Stress, Solid Int Prin Stress -->  Solid Triaxial Stress

		rc = feoutput.Get(300055)

		ReDim c1(nElem) As Double
		ReDim c2(nElem) As Double
		ReDim c3(nElem) As Double
		ReDim c4(nElem) As Double
		ReDim c5(nElem) As Double
		ReDim c6(nElem) As Double
		ReDim c7(nElem) As Double
		ReDim c8(nElem) As Double
		ReDim cenVal1(nElem) As Double

		For j = 0 To nElem-1
			c1(j) = smaxpc1(j) + sminpc1(j) + sintpc1(j)
			c2(j) = smaxpc2(j) + sminpc2(j) + sintpc2(j)
			c3(j) = smaxpc3(j) + sminpc3(j) + sintpc3(j)
			c4(j) = smaxpc4(j) + sminpc4(j) + sintpc4(j)
			c5(j) = smaxpc5(j) + sminpc5(j) + sintpc5(j)
			c6(j) = smaxpc6(j) + sminpc6(j) + sintpc6(j)
			c7(j) = smaxpc7(j) + sminpc7(j) + sintpc7(j)
			c8(j) = smaxpc8(j) + sminpc8(j) + sintpc8(j)
			cenVal1(j) = smaxpc(j) + sminpc(j) + sintpc(j)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitElemWithCorner(OutputSetID, 300055, 171, 172, 173, 174, 175, 176, 177, 178,"Solid Triaxial Stress",0,True)
		rc = feoutput.PutElemWithCorner(nElem, maxCorner, eID, cenVal1, c1, c2, c3, c4, c5, c6, c7, c8 )
		rc = feoutput.Put(300055)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		rc = App.feAppMessage( 0, "Output Vectors of Solids for Output case" + Str(OutputSetID) + " are done")
	Else
		rc = App.feAppMessage( 0, "There are no Output Vectors of Solids for Output case" + Str(OutputSetID))
	End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                                                                    B E A M    S T R E S S E S
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 3022/Area-->3020400
' Beam EndA Axial Force, Beam Area --> Beam EndA Axial Stress
	Dim pr As femap.Prop
	Set pr = App.feProp

	Dim El As femap.Elem
	Set El = App.feElem

	feoutput.setID = OutputSetID

	If feoutput.Get(3022) Then
		vout = feoutput.vcomponent
		rc = feoutput.GetScalarAtElem(nElem,eID,p)

		ReDim asp(nElem) As Double
		ReDim A(nElem) As Double

		For j = 0 To nElem-1
			rc = El.Get(eID(j))
			propID = El.propID
			rc = pr.Get(propID)
			A(j) = pr.pval(0)
			asp(j) = p(j)/A(j)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitScalarAtElem(OutputSetID,302400,"Beam EndA Axial Stress",0,False)
		rc = feoutput.PutScalarAtElem(nElem,eID,asp)
		vout(0) = 302400
		vout(1) = 302500
		feoutput.hascomponent = 3
		feoutput.vcomponent = vout
		rc = feoutput.Put(302400)

' 3023/Area-->3020500
' Beam EndB Axial Force, Beam Area --> Beam EndB Axial Stress
		rc = feoutput.Get(3023)
		vout = feoutput.vcomponent
		rc = feoutput.GetScalarAtElem(nElem,eID,p)

		ReDim asp(nElem) As Double
		ReDim A(nElem) As Double

		For j = 0 To nElem-1
			rc = El.Get(eID(j))
			propID = El.propID
			rc = pr.Get(propID)
			A(j) = pr.pval(0)
			asp(j) = p(j)/A(j)

			If status > bigstatus Then
   				App.feAppStatusUpdate (status)
   				App.feAppStatusRedraw()
				bigstatus = bigstatus + 1000
			End If
			status = status + 1

		Next j

		rc = feoutput.InitScalarAtElem(OutputSetID,302500,"Beam EndB Axial Stress",0,False)
		rc = feoutput.PutScalarAtElem(nElem,eID,asp)
		vout(0) = 302400
		vout(1) = 302500
		feoutput.hascomponent = 3
		feoutput.vcomponent = vout
		rc = feoutput.Put(302500)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		rc = App.feAppMessage( 0, "Output Vectors of Beams for Output case" + Str(OutputSetID) + " are done")
	Else
		rc = App.feAppMessage( 0, "There are no Output Vectors of Beams for Output case" + Str(OutputSetID))
	End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



	OutputSetID = os.Next
Wend
rc = App.feAppMessage(0,"Program Completed")
App.feAppStatusShow(False, 1)
App.feAppStatusRedraw()
End Sub
