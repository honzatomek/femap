Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long, k As Long

	Dim LCSet As Object
	Set LCSet = App.feSet()

	Dim LC() As Object

	Dim vec As Object
	Set vec = App.feOutput()

	Dim vecIDs As Variant

	vecIDs = Array(192, 193, 194, 196, 197, 198)

	Dim ndset As Object
	Set ndset = App.feSet()

	Dim ndIDs() As Long
	Dim forces() As Variant

	Dim modelID1 As PortInt
	Dim modelID2 As PortInt

	Dim numModels As Long
    Dim modelIDs As Variant

    rc = App.feAppGetAllModels(numModels, modelIDs)

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

	modelID1 = selectmodel(0, "Select Model to copy LCs from:")
	If modelID1 = -1 Then
		er = 1
		GoTo Cleanup
	End If

	rc = 0
	Do While rc <> -1
		modelID2 = selectmodel(1, "Select Model to copy  LCs to:")
		If modelID2 = -1 Then
			er = 2
			GoTo Cleanup
		ElseIf modelID1 <> modelID2 Then
			rc = -1
		End If
	Loop

	modelID1 = modelIDs(modelID1)
	modelID2 = modelIDs(modelID2)

	rc = App.feAppSetModel(modelID1)

	rc = LCSet.SelectMultiID(FT_OUT_CASE, 1, "Select Output Sets To be Transformed:")
	If (rc <> -1) Or (LCSet.Count() < 1) Then
		er = 3
		GoTo Cleanup
	End If

	rc = ndset.Select(FT_NODE, True, "Select Nodes for Trasformation:")
	If (rc <> -1) Or (ndset.Count() < 1) Then
		er = 4
		GoTo Cleanup
	End If

	ReDim ndIDs(ndset.Count() - 1)
	ReDim forces(ndset.Count() - 1)

	rc = ndset.Reset()
	For i = 0 To ndset.Count() - 1 Step 1
		ndIDs(i) = ndset.Next()
		forces(i) = Array(0,0,0,0,0,0)
	Next

	Dim LCnew As Object
	Set LCnew = App.feLoadSet()

	Dim Ld As Object
	Set Ld = App.feLoadMesh()

	ReDim LC(LCSet.Count() - 1)
	rc = LCSet.Reset()
	For i = 0 To LCSet.Count() - 1 Step 1
		rc = App.feAppSetModel(modelID1)
		Set LC(i) = App.feOutputSet()
		rc = LC(i).Get(LCSet.Next())
		For j = 0 To 5 Step 1
			Set vec = LC(i).Vector(vecIDs(j))
			For k = 0 To UBound(ndIDs) Step 1
				forces(k)(j) = vec.Value(ndIDs(k))
			Next
		Next

		rc = App.feAppSetModel(modelID2)
		LCnew.ID = LCnew.NextEmptyID()
		LCnew.title = LC(i).title
		rc = LCnew.Put(LCnew.ID)
		Ld.CSys = 0
		Ld.setID = LCnew.ID
		Ld.XOn = True
		Ld.YOn = True
		Ld.ZOn = True
		For j = 0 To UBound(ndIDs)
			Ld.ID = Ld.NextEmptyID()
			Ld.meshID = ndIDs(j)
			Ld.type = 1 '1 = nForce, 2 = nMoment
			Ld.load(0) = forces(j)(0)
			Ld.load(1) = forces(j)(1)
			Ld.load(2) = forces(j)(2)
			rc = Ld.Put(Ld.ID)
		Next
	Next

	er = 0
Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub

Private Function selectmodel(Optional ByVal preselect As Long = 0, Optional ByVal caption As String = "Select Model:") As Long
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

	Dim ListArray() As String

	Dim numModels As Long
    Dim modelIDs As Variant

    rc = App.feAppGetAllModels(numModels, modelIDs)

    ReDim ListArray(numModels - 1)

    For i = 0 To numModels - 1 Step 1
		rc = App.feAppGetModelName(modelIDs(i), ListArray(i))
		If InStr(1, ListArray(i), "\") <> 0 Then
			ListArray(i) = Mid$(ListArray(i), InStrRev(ListArray(i), "\") + 1)
		End If
		ListArray(i) = i + 1 & " - " & ListArray(i)
    Next

	Begin Dialog UserDialog 550,189,caption,.SelectModelFunc 'Allows User To Select Model from a list of Opened Models %GRID:10,7,1,1
		ListBox 10,35,530,119,ListArray(),.ListBox1
		OKButton 360,161,90,21
		CancelButton 460,161,80,21
		TextBox 10,7,530,21,.TextBox1
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox1 = ListArray(preselect)
	dlg.ListBox1 = preselect
	txt = ListArray(preselect)
	rc = Dialog(dlg)

	If rc <> -1 Then
		selectmodel = -1
	Else
		selectmodel = dlg.ListBox1
	End If

On Error Resume Next
	Set App = Nothing
End Function

Private txt As String

Rem See DialogFunc help topic for more information.
Private Function selectmodelfunc(DlgItem$, Action%, SuppValue?) As Boolean
	Dim tmp
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem selectmodel = True ' Prevent button press from closing the dialog box
		selectmodelfunc = True
		Select Case DlgItem$
		Case "OK"
			If IsNumeric(DlgText("TextBox1")) Then
				tmp = CLng(DlgText("TextBox1"))
				If (tmp > 0) And (tmp < DlgListBoxArray("ListBox1") + 1) Then
					Call DlgValue("ListBox1", tmp - 1)
					Call DlgText("TextBox1", DlgText("ListBox1"))
					txt = DlgText("TextBox1")
					selectmodelfunc = False
				Else
					DlgText("TextBox1", txt)
				End If
			Else
				If DlgText("TextBox1") =DlgText("ListBox1") Then
					selectmodelfunc = False
				Else
					tmp =DlgValue("ListBox1")
					Call DlgText("TextBox1", DlgText("ListBox1"))
				End If
			End If
		Case "Cancel"
			selectmodelfunc = False
		Case "ListBox1"
			Call DlgText("TextBox1", DlgText(DlgItem))
			txt = DlgText("TextBox1")
		Case "TextBox1"
		End Select
	Case 3 ' TextBox or ComboBox text changed
		Select Case DlgItem$
		Case "TextBox1"

		End Select
	Case 4 ' Focus changed
		If SuppValue? < DlgCount() Then
			Select Case DlgName(SuppValue?)
			Case "TextBox1"
				If IsNumeric(DlgText("TextBox1")) Then
					tmp = CLng(DlgText("TextBox1"))
					If (tmp > 0) And (tmp < DlgListBoxArray("ListBox1") + 1) Then
						Call DlgValue("ListBox1", tmp - 1)
						Call DlgText("TextBox1", DlgText("ListBox1"))
						txt = DlgText("TextBox1")
					Else
						DlgText("TextBox1", txt)
					End If
				Else
					tmp =DlgValue("ListBox1")
					Call DlgText("TextBox1", DlgText("ListBox1"))
				End If
			End Select
			Select Case DlgItem
			Case "TextBox1"
			End Select
		End If
	Case 5 ' Idle
		Rem Wait .1 : selectmodel = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
