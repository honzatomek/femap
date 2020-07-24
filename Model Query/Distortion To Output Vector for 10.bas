Rem File: DistortionToOutputVector.Bas
Sub Main
    Dim App As femap.model
    Set App = feFemap()

'------------------------------
Const maxEls = 100000

'------------------------------
Dim ol As femap.Set
Set ol = App.feSet

Dim rc As Long
Dim el As femap.Elem
Set el = App.feElem

Dim Eid As Long

Dim elLIST(maxEls) As Long
Dim aspect(maxEls) As Double
Dim angle(maxEls) As Double
Dim Taper(maxEls) As Double
Dim warp(maxEls) As Double
Dim alttaper(maxEls) As Double
Dim tet(maxEls) As Double
Dim Jacob(maxEls) As Double
Dim Comb(maxEls) As Double
Dim NasWarp(maxEls) As Double

Dim status As Long
Dim count As Long

rc = ol.Select(8, True, "Choose Elements for Distortion Plot")

If rc = -1 Then

count = ol.count()
If Not (count < maxEls) Then
    MsgBox ("Error, You must select less than" + Str$(maxEls) +" Elements at a time.")
End If

If (count < maxEls) Then

count = count + 3
status = 1

J = App.feAppStatusShow(True, count)
App.feAppStatusUpdate (status)
J = App.feAppStatusRedraw()

rc = App.feAppLock()

Eid = ol.Next()
i = 0

Do While Eid > 0
    j = App.feGetElemDistortion(Eid, aspect(i), Taper(i), angle(i), warp(i), NasWarp(i),  alttaper(i), tet(i), Jacob(i), Comb(i) )
   
   status = status + 1
   App.feAppStatusUpdate (status)
   
    elLIST(i) = Eid
    i = i + 1
    Eid = ol.Next()
   
Loop

rc = App.feAppUnlock()

j = App.feAppStatusRedraw()
    
Dim output0 As femap.output
Set output0 = App.feOutput
Dim output1 As femap.output
Set output1 = App.feOutput
Dim output2 As femap.output
Set output2 = App.feOutput
Dim output3 As femap.output
Set output3 = App.feOutput
Dim output4 As femap.output
Set output4 = App.feOutput
Dim output5 As femap.output
Set output5 = App.feOutput
Dim output6 As femap.output
Set output6 = App.feOutput
Dim output7 As femap.output
Set output7 = App.feOutput
Dim output8 As femap.output
Set output8 = App.feOutput

Dim setID As Long
Dim outset As femap.OutputSet
Set outset = App.feOutputSet

Dim elIDV As Variant
Dim aspectV As Variant
Dim angleV As Variant
Dim TaperV As Variant
Dim warpV As Variant
Dim alttaperV As Variant
Dim tetV As Variant
Dim JacV As Variant
Dim NWarpV As Variant
Dim CombV As Variant

    setID = outset.NextEmptyID()

    'Create the output set
    outset.title = "Distortion Set"
    outset.value = 0
    outset.analysis = 0
    outset.Put (setID)
            
    rc = output0.InitScalarAtElem(setID, 400000, "Aspect Ratio", 4, True)
    rc = output1.InitScalarAtElem(setID, 400001, "Taper", 4, True)
    rc = output2.InitScalarAtElem(setID, 400002, "Alt Taper", 4, True)
    rc = output3.InitScalarAtElem(setID, 400003, "Int Angles", 4, True)
    rc = output4.InitScalarAtElem(setID, 400004, "Warping", 4, True)
    rc = output5.InitScalarAtElem(setID, 400005, "Tet Collapse", 4, True)
	rc = output6.InitScalarAtElem(setID, 400006, "Jacobian", 4, True)
	rc = output7.InitScalarAtElem(setID, 400007, "Nastran Warping", 4, True)
    rc = output8.InitScalarAtElem(setID, 400008, "Combined Quality", 4, True)

	k = i
   
    status = status + 1
    App.feAppStatusUpdate (status)
    j = App.feAppStatusRedraw()
    

	elIDV = elLIST
	aspectV = aspect
	angleV = angle
	TaperV = Taper
	warpV = warp	
	alttaperV = alttaper
	tetV = tet
	JacV = Jacob
	NWarpV = NasWarp
	CombV = Comb

	rc = output0.PutScalarAtElem(k, elIDV, aspectV)
	rc = output0.Put(-1)
	rc = output1.PutScalarAtElem(k, elIDV, TaperV)
	rc = output1.Put(-1)
	rc = output2.PutScalarAtElem(k, elIDV, alttaperV)
	rc = output2.Put(-1)
	rc = output3.PutScalarAtElem(k, elIDV, angleV)
	rc = output3.Put(-1)
	rc = output4.PutScalarAtElem(k, elIDV, warpV)
	rc = output4.Put(-1)
	rc = output5.PutScalarAtElem(k, elIDV, tetV)
	rc = output5.Put(-1)
	rc = output6.PutScalarAtElem(k, elIDV, JacV)
	rc = output6.Put(-1)
	rc = output7.PutScalarAtElem(k, elIDV, NWarpV)
	rc = output7.Put(-1)
	rc = output8.PutScalarAtElem(k, elIDV, CombV)
	rc = output8.Put(-1)

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
    Str = " created for Distortion data"
    sval = sval + Str
    
    k = App.feAppMessage(color, sval)
        
    End If
End If

j = App.feAppStatusShow(False, 4)
j = App.feAppStatusRedraw()
    
End Sub


