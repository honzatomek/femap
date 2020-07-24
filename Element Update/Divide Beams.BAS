Rem File: dividebeams.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim Prop1 As femap.Prop
	Set Prop1 = App.feProp

	Dim El As femap.Elem
	Set El = App.feElem

	Dim Node As femap.Node
	Set Node = App.feNode

	Dim PID As Long
	Dim PIDW As Long

	Dim EID As Long
	Dim EIDW As Long

	Dim nd1 As Long
	Dim nd2 As Long
	Dim nd1X As Double
	Dim nd2X As Double
	Dim nd1Y As Double
	Dim nd2Y As Double
    Dim nd1Z As Double
	Dim nd2Z As Double
	Dim distx As Double
	Dim disty As Double
	Dim distz As Double

'Beam Section Type

    Dim BType As Integer

'Property Values A

	Dim AreaA As Double
	Dim IZZA As Double
	Dim IYYA As Double
	Dim IZYA As Double
	Dim JA As Double
	Dim YSHEARA As Double
	Dim ZSHEARA As Double
	Dim NSMA As Double
	Dim WARPA As Double
	Dim YOFFA As Double
	Dim ZOFFA As Double

'Stress Recovery Values A

	Dim YA1 As Double
	Dim YA2 As Double
	Dim YA3 As Double
	Dim YA4 As Double

	Dim ZA1 As Double
	Dim ZA2 As Double
	Dim ZA3 As Double
	Dim ZA4 As Double

'Section Values A

    Dim HA As Double
	Dim RADA As Double
	Dim W1A As Double
	Dim W2A As Double
	Dim T1A As Double
	Dim T2A As Double
	Dim TA As Double

'Property Values B

	Dim AreaB As Double
	Dim IZZB As Double
	Dim IYYB As Double
	Dim IZYB As Double
	Dim JB As Double
	Dim YSHEARB As Double
	Dim ZSHEARB As Double
	Dim NSMB As Double
	Dim WARPB As Double
	Dim YOFFB As Double
	Dim ZOFFB As Double

'Stress Recovery Values B

	Dim YB1 As Double
	Dim YB2 As Double
	Dim YB3 As Double
	Dim YB4 As Double

	Dim ZB1 As Double
	Dim ZB2 As Double
	Dim ZB3 As Double
	Dim ZB4 As Double

'Section Values B

    Dim HB As Double
	Dim RADB As Double
	Dim W1B As Double
	Dim W2B As Double
	Dim T1B As Double
	Dim T2B As Double
	Dim TB As Double

'Property Values Working

	Dim AreaW As Double
	Dim IZZW As Double
	Dim IYYW As Double
	Dim IZYW As Double
	Dim JW As Double
	Dim YSHEARW As Double
	Dim ZSHEARW As Double
	Dim NSMW As Double
	Dim WARPW As Double
	Dim YOFFW As Double
	Dim ZOFFW As Double

'Stress Recovery Values Working

	Dim YW1 As Double
	Dim YW2 As Double
	Dim YW3 As Double
	Dim YW4 As Double

	Dim ZW1 As Double
	Dim ZW2 As Double
	Dim ZW3 As Double
	Dim ZW4 As Double

'Section Values Working

    Dim HW As Double
	Dim RADW As Double
	Dim W1W As Double
	Dim W2W As Double
	Dim T1W As Double
	Dim T2W As Double
	Dim TW As Double

'Select a Property

	Dim pset As femap.Set
	Set pset = App.feSet

SelectBeam:
rc = pset.SelectID (FT_ELEM, "Select Element", EID)

If rc=FE_CANCEL Then
    GoTo FAIL
    End If


	El.Get (EID)
	If El.Prop.type <> 5 And El.Prop.type <> 36 Then
		App.feAppMessage ( FCM_ERROR, "Selected Element Must be a Beam - Reselect" )
		GoTo SelectBeam
	End If

	Begin Dialog UserDialog 400,126,"Number of Divisions"
		text 40,21,130,21,"Number of Divisions",.Text1
		TextBox 210,21,140,21,.div
		OKButton 90,98,90,21
		CancelButton 230,98,90,21
	End Dialog
	Dim dlg As UserDialog
	'Dialog dlg

	If Dialog(dlg) = 0 Then
	GoTo FAIL
	End If

	Dim Div As Long
	Dim Divs As Long

	Dim Per As Double
	Divs = (Val (dlg.div))

	nd1 = El.Node(0)

	Node.Get (nd1)

	nd1X = Node.x
   	nd1Y = Node.y
   	nd1Z = Node.z

   	nd2 = El.Node(1)

  Node.Get(nd2)

   nd2X = Node.x
   nd2Y = Node.y
   nd2Z = Node.z

   PID = Prop1.Get (El.propID)

    HA = Prop1.pval (40)
    RADA = Prop1.pval (41)
	W1A = Prop1.pval (42)
    W2A = Prop1.pval (43)
    T1A = Prop1.pval (44)
	T2A = Prop1.pval (45)
	TA = Prop1.pval (46)

For J = 1 To (Divs-1) Step 1

Div = J

Per = (Div/Divs)

	If Prop1.flag(0) = 0 Then

	PIDW = Prop1.ID

	GoTo Divide

	End If

	AreaB = Prop1.pval (20)
	IZZB = Prop1.pval(21)
	IYYB = Prop1.pval(22)
	IZYB = Prop1.pval(23)
	JB = Prop1.pval(24)
	YSHEARB = Prop1.pval(25)
	ZSHEARB = Prop1.pval(26)
	NSMB = Prop1.pval(27)
	WARPB = Prop1.pval(38)
	YOFFB = Prop1.pval(36)
	ZOFFB = Prop1.pval(37)

	YB1 = Prop1.pval(28)
	YB2 = Prop1.pval(30)
	YB3 = Prop1.pval(32)
	YB4 = Prop1.pval(34)

	ZB1 = Prop1.pval(29)
	ZB2 = Prop1.pval(31)
	ZB3 = Prop1.pval(33)
	ZB4 = Prop1.pval(35)

	BType = Prop1.flagI(1)

	HB = Prop1.pval (55)
    RADB = Prop1.pval (56)
	W1B = Prop1.pval (57)
    W2B = Prop1.pval (58)
    T1B = Prop1.pval (59)
	T2B = Prop1.pval (60)
	TB = Prop1.pval (61)

	HW = HA - ((HA-HB)*Per)
    RADW = RADA - ((RADA-RADB)*Per)
	W1W = W1A - ((W1A-W1B)*Per)
    W2W = W2A - ((W2A-W2B)*Per)
    T1W = T1A - ((T1A-T1B)*Per)
	T2W = T2A - ((T2A-T2B)*Per)
	TW = TA-((TA-TB)*Per)

	Prop1.flagI(1) = BType
    Prop1.pval (55) = HW
    Prop1.pval (56) = RADW
    Prop1.pval (57) = W1W
    Prop1.pval (58) = W2W
    Prop1.pval (59) = T1W
	Prop1.pval (60) = T2W
	Prop1.pval (61) = TW

	Dim SHEARA As Boolean

	SHEARA = True

	If YOFFB = 0.0 Then

		If ZOFFB = 0.0 Then

		SHEARA = False

		End If

	End If

	Prop1.ComputeShape (SHEARA, False, True)

	Prop1.Put (PID)

'Get values from End B in new property and make them end A

	Prop1.Get (PID)

	AreaW = Prop1.pval (20)
	IZZW = Prop1.pval(21)
	IYYW = Prop1.pval(22)
	IZYW = Prop1.pval(23)
	JW = Prop1.pval(24)
	YSHEARW = Prop1.pval(25)
	ZSHEARW = Prop1.pval(26)
	NSMW = Prop1.pval(27)
	WARPW = Prop1.pval(38)
	YOFFW = Prop1.pval(36)
	ZOFFW = Prop1.pval(37)

	YW1 = Prop1.pval(28)
	YW2 = Prop1.pval(30)
	YW3 = Prop1.pval(32)
	YW4 = Prop1.pval(34)

	ZW1 = Prop1.pval(29)
	ZW2 = Prop1.pval(31)
	ZW3 = Prop1.pval(33)
	ZW4 = Prop1.pval(35)

    HW = Prop1.pval (55)
    RADW = Prop1.pval (56)
	W1W = Prop1.pval (57)
    W2W = Prop1.pval (58)
    T1W = Prop1.pval (59)
	T2W = Prop1.pval (60)
	TW = Prop1.pval (61)

	Prop1.pval (0) = AreaW
	Prop1.pval(1) = IZZW
	Prop1.pval(2) = IYYW
	Prop1.pval(3) = IZYW
	Prop1.pval(4) = JW
	Prop1.pval(5) = YSHEARW
	Prop1.pval(6) = ZSHEARW
	Prop1.pval(7) = NSMW
	Prop1.pval(18) = WARPW
	Prop1.pval(16) = YOFFW
	Prop1.pval(17) = ZOFFW

	Prop1.pval(8) = YW1
	Prop1.pval(10) = YW2
	Prop1.pval(12) = YW3
	Prop1.pval(14) = YW4

	 Prop1.pval(9) = ZW1
	 Prop1.pval(11) = ZW2
	 Prop1.pval(13) = ZW3
	 Prop1.pval(15) = ZW4

	Prop1.flagI(1) = BType
    Prop1.pval (40) = HW
    Prop1.pval (41) = RADW
    Prop1.pval (42) = W1W
    Prop1.pval (43) = W2W
    Prop1.pval (44) = T1W
	Prop1.pval (45) = T2W
	Prop1.pval (46) = TW

    'Prop1.ComputeShape (False, False, True)

	Prop1.pval (20) = AreaB
	Prop1.pval(21) = IZZB
	Prop1.pval(22) = IYYB
	Prop1.pval(23) = IZYB
	Prop1.pval(24) = JB
	Prop1.pval(25) = YSHEARB
	Prop1.pval(26) = ZSHEARB
	Prop1.pval(27) = NSMB
	Prop1.pval(38) = WARPB
	Prop1.pval(36) = YOFFB
	Prop1.pval(37) = ZOFFB

	Prop1.pval(28) = YB1
	Prop1.pval(30) = YB2
	Prop1.pval(32) = YB3
	Prop1.pval(34) = YB4

	 Prop1.pval(29) = ZB1
	 Prop1.pval(31) = ZB2
	 Prop1.pval(33) = ZB3
	 Prop1.pval(35) = ZB4

	Prop1.flagI(1) = BType
    Prop1.pval (55) = HB
    Prop1.pval (56) = RADB
    Prop1.pval (57) = W1B
    Prop1.pval (58) = W2B
    Prop1.pval (59) = T1B
	Prop1.pval (60) = T2B
	Prop1.pval (61) = TB

	'Prop1.ComputeShape (False, False, True)

	 PIDW = Prop1.NextEmptyID

	 Prop1.title = "Beam Section"

     Prop1.Put (PIDW)

     Divide:

   distx = (nd1X + (nd2X - nd1X) * Per)
   disty = (nd1Y + (nd2Y - nd1Y) * Per)
   distz = (nd1Z + (nd2Z - nd1Z) * Per)

  ndw = Node.NextEmptyID

   Node.x = distx
   Node.y = disty
   Node.z = distz

   Node.Put (ndw)

   El.Node(1) = ndw

   El.Put (EID)

	EIDW = El.NextEmptyID

	El.Get (EIDW)

	El.Node (0) = ndw
	El.Node (1) = nd2

	El.propID = PIDW

	El.Put (EIDW)

	EID = EIDW

Next J

App.feViewRegenerate ( 0 )
GoTo SelectBeam

	FAIL:

End Sub
