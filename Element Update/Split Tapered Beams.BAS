Rem File: splittaperbeam.BAS
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

'Section Values A

    Dim HA As Double
	Dim RADA As Double
	Dim W1A As Double
	Dim W2A As Double
	Dim T1A As Double
	Dim T2A As Double
	Dim TA As Double

'Section Values B

    Dim HB As Double
	Dim RADB As Double
	Dim W1B As Double
	Dim W2B As Double
	Dim T1B As Double
	Dim T2B As Double
	Dim TB As Double

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

rc = pset.SelectID (FT_ELEM, "Select Element", EID)

	Begin Dialog UserDialog 400,126,"Enter Split Distance"
		text 40,21,130,21,"Split Distance (%)",.Text1
		TextBox 210,21,140,21,.per
		OKButton 90,98,90,21
		CancelButton 230,98,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

	Dim Per As Double
	Per = (Val (dlg.per))*.01

	El.Get (pset.First)

	PID = Prop1.Get (El.propID)

	If Prop1.flag(0) = 0 Then

	PIDW = Prop1.ID

	GoTo Divide

	End If

	BType = Prop1.flagI(1)

	YOFFB = Prop1.pval(36)
	ZOFFB = Prop1.pval(37)

    HA = Prop1.pval (40)
    RADA = Prop1.pval (41)
	W1A = Prop1.pval (42)
    W2A = Prop1.pval (43)
    T1A = Prop1.pval (44)
	T2A = Prop1.pval (45)
	TA = Prop1.pval (46)

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

    HW = Prop1.pval (55)
    RADW = Prop1.pval (56)
	W1W = Prop1.pval (57)
    W2W = Prop1.pval (58)
    T1W = Prop1.pval (59)
	T2W = Prop1.pval (60)
	TW = Prop1.pval (61)

	Prop1.flagI(1) = BType
    Prop1.pval (40) = HW
    Prop1.pval (41) = RADW
    Prop1.pval (42) = W1W
    Prop1.pval (43) = W2W
    Prop1.pval (44) = T1W
	Prop1.pval (45) = T2W
	Prop1.pval (46) = TW

    Prop1.ComputeShape (SHEARA, False, True)

	Prop1.flagI(1) = BType
    Prop1.pval (55) = HB
    Prop1.pval (56) = RADB
    Prop1.pval (57) = W1B
    Prop1.pval (58) = W2B
    Prop1.pval (59) = T1B
	Prop1.pval (60) = T2B
	Prop1.pval (61) = TB

	Prop1.ComputeShape (SHEARA, False, True)

	 PIDW = Prop1.NextEmptyID

	 Prop1.title = "Beam Section"

     Prop1.Put (PIDW)

     Divide:

     El.Get (EID)

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

	App.feViewRegenerate (0)

End Sub
