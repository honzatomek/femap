Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long
    Dim sumFC() As Variant
    Dim sumMS() As Variant

    er = 0

    Dim useBodyLoad As Boolean
    Dim expandGEOM As Boolean
    Dim doLIST As Boolean
    Dim useSETS As Boolean
    Dim nodeSET As Long
    Dim elemSET As Long
    Dim loaddefSET As Long
    Dim basePOINT(2) As Double
    Dim csysID As Long
    Dim summedFORCES As Variant

    useBodyLoad = True
    expandGEOM = True
    doLIST = True
    useSETS = False
    nodeSET = 0
    elemSET = 0
    loaddefSET = 0
    basePOINT(0) = 0
    basePOINT(1) = 0
    basePOINT(2) = 0
    csysID = 0

    Dim ldsets As Object
    Set ldsets = App.feSet()
    rc = ldsets.AddAll(FT_LOAD_DIR)

    Dim LdSet As Object
    Set LdSet = App.feLoadSet()

    Dim activeLD As Long

    activeLD = App.Info_ActiveID(FT_LOAD_DIR)

    Dim eSET As Object
    Dim printResults As Boolean
    Dim printDetails As Boolean
    Dim length As Double
    Dim area As Double
    Dim volume As Double
    Dim structMass As Double
    Dim nonstructMass As Double
    Dim totalMass As Double
    Dim structCG As Variant
    Dim nonstructCG As Variant
    Dim totalCG As Variant
    Dim inertia As Variant
    Dim inertiaCG As Variant

    Set eSET = App.feSet()
    rc = eSET.AddAll(FT_ELEM)
    printResults = True
    printDetails = False

    If ldsets.Count() = 0 Then GoTo Cleanup

    ReDim sumFC(0)
    sumFC(0) = Array("Load Case", "Fx [daN]", "Fy [daN]", "Fz [daN]")
    ReDim sumMS(0)
    sumMS(0) = Array("Load Case", "Mass [kg]")

	Call App.feAppMessageClear()
    rc = ldsets.Reset()
    Do While ldsets.Next() <> 0
    	rc = LdSet.Get(ldsets.CurrentID())
		App.Info_ActiveID(FT_LOAD_DIR) = LdSet.ID()
		Call App.feAppMessage(FCM_COMMAND, "[" & LdSet.ID() & "].." & LdSet.title())
		Call App.feAppMessage(FCM_WARNING, "Check Sum Of Forces")
		rc = App.feCheckSumForces2(useBodyLoad	, expandGEOM, doLIST, useSETS, nodeSET, elemSET, loaddefSET, basePOINT, csysID, summedFORCES)
		Call App.feAppMessage(FCM_COMMAND, "Total: Fx = " & Replace$(Format$(summedFORCES(9) / 10, "# ### ##0.0 daN"), ",", ".") & ", Fy = " & Replace$(Format$(summedFORCES(10) / 10, "# ### ##0.0 daN"), ",", ".") & ", Fz = " & Replace$(Format$(summedFORCES(11) / 10, "# ### ##0.0 daN"), ",", "."))
		Call App.feAppMessage(FCM_NORMAL, "")
		ReDim Preserve sumFC(UBound(sumFC) + 1)
		sumFC(UBound(sumFC)) = Array("[" & LdSet.ID() & "].." & LdSet.title(), Replace$(Format$(summedFORCES(9) / 10, "# ### ##0.0"), ",", "."), Replace$(Format$(summedFORCES(10) / 10, "# ### ##0.0"), ",", "."), Replace$(Format$(summedFORCES(11) / 10, "# ### ##0.0"), ",", "."))
		If LdSet.BodyAccelOn() Then
			Call App.feAppMessage(FCM_WARNING, "Mass Properties")
			rc = App.feMeasureMeshMassProp(eSET.ID(), csysID, printResults, printDetails, length, area, volume, structMass, nonstructMass, totalMass, structCG, nonstructCG, totalCG,  inertia,  inertiaCG)
			Call App.feAppMessage(FCM_COMMAND, "Total: " &  Replace$(Format$(totalMass * 1000, "# ### ##0.0 kg"), ",", "."))
			Call App.feAppMessage(FCM_NORMAL, "")
			ReDim Preserve sumMS(UBound(sumMS) + 1)
			sumMS(UBound(sumMS)) = Array("[" & LdSet.ID() & "].." & LdSet.title(), Replace$(Format$(totalMass * 1000, "# ### ##0.0"), ",", "."))
		End If
    Loop

	Dim col1 As Long, col2 As Long
	col1 = MaxLen(sumFC, 0) + 1
	col2 = Max(Array(MaxLen(sumFC, 1), MaxLen(sumFC, 2), MaxLen(sumFC, 3))) + 1
	For i = LBound(sumFC) To UBound(sumFC)
		sumFC(i)(0) = LeftAlign(sumFC(i)(0), col1)
		For j = LBound(sumFC(i)) + 1 To UBound(sumFC(i))
			sumFC(i)(j) = RightAlign(sumFC(i)(j), col2)
		Next
	Next
    Call App.feAppMessage(FCM_NORMAL, "")
    Call App.feAppMessage(FCM_COMMAND, "Forces Overview:")
    Call App.feAppMessage(FCM_NORMAL, "")
    Call App.feAppMessage(FCM_NORMAL, Join(sumFC(0), " "))
    For i = LBound(sumFC) + 1 To UBound(sumFC) Step 1
		Call App.feAppMessage(FCM_NORMAL, Join(sumFC(i), " "))
    Next

	If UBound(sumMS) > 0 Then
		col1 = MaxLen(sumMS, 0) + 1
		col2 = MaxLen(sumMS, 1) + 1
		For i = LBound(sumMS) To UBound(sumMS)
			sumMS(i)(0) = LeftAlign(sumMS(i)(0), col1)
			For j = LBound(sumMS(i)) + 1 To UBound(sumMS(i))
				sumMS(i)(j) = RightAlign(sumMS(i)(j), col2)
			Next
		Next
	    Call App.feAppMessage(FCM_NORMAL, "")
	    Call App.feAppMessage(FCM_COMMAND, "Mass Overview:")
	    Call App.feAppMessage(FCM_NORMAL, "")
	    Call App.feAppMessage(FCM_NORMAL, Join(sumMS(0), " "))
	    For i = LBound(sumMS) + 1 To UBound(sumMS) Step 1
			Call App.feAppMessage(FCM_NORMAL, Join(sumMS(i), " "))
	    Next
	End If

	App.Info_ActiveID(FT_LOAD_DIR) = activeLD

Cleanup:
	If DBUG Then Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set eSET = Nothing
	Set LdSet = Nothing
	Set ldsets = Nothing
	Set App = Nothing
End Sub

Function LeftAlign(ByVal txt As String, ByVal length As Long) As String
	txt = Left(txt, length - 1)
	LeftAlign = txt & String(length - Len(txt), " ")
End Function

Function RightAlign(ByVal txt As String, ByVal length As Long) As String
	txt = Left(txt, length - 1)
	RightAlign =String(length - Len(txt), " ") & txt
End Function

Function MaxLen(ByRef arr As Variant, ByVal col As Long) As Long
	Dim i As Long
	MaxLen = 0
	For i = LBound(arr) To UBound(arr)
		If Len(arr(i)(col)) > MaxLen Then MaxLen = Len(arr(i)(col))
	Next
End Function

Function Max(ByRef vals As Variant) As Long
	Dim i As Long
	Max = vals(LBound(vals))
	For i = LBound(vals) + 1 To UBound(vals)
		If vals(i) > Max Then Max = vals(i)
	Next
End Function
