Option Explicit On

Const DBUG As Boolean = False
Private er As Long
Private ldir As Long

'load elements by function with variables x0, y0, z0, x1, y1, z1, A, L

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim elset As Object
    Set elset = App.feSet()

    Dim el As Object
    Set el = App.feElem()

    Dim nd As Object
    Set nd = App.feNode()

    Dim vecn(8) As String
    Dim v(8) As Double

    Dim A0 As Double
    Dim A1 As Double
    Dim L As Double

    vecn(0) = "X0"
    vecn(1) = "Y0"
    vecn(2) = "Z0"
    vecn(3) = "X1"
    vecn(4) = "Y1"
    vecn(5) = "Z1"
    vecn(6) = "A0"
    vecn(7) = "A1"
    vecn(8) = "L"

    Dim load As Double
    Dim ld As Object
    Set ld = App.feLoadMesh()

    Dim DOF(2) As Long
    DOF(0) =6
    DOF(1) = 0
    DOF(2) = 0

	Dim vals(5) As Double
	For i = 0 To 5 Step 1
		vals(i) = 0
	Next
	Dim func(4) As Long
	For i = 0 To 4 Step 1
		func(i) = 0
	Next

	Do
	    rc = elset.Select(8, True, "Select elements to load:")
		If (rc <> -1) Or (elset.Count() < 1) Then
			er = 1
			GoTo Cleanup
		End If

		Dim min As Double
		Dim max As Double
		Dim Length As Double

		rc = elset.Reset()
		rc = el.Get(elset.First())
		rc = nd.Get(el.Node(0))
		max = nd.x
		rc = nd.Get(el.Node(1))
		min = nd.x

		rc = elset.Reset()
		Do While elset.Next() <> 0
			rc = el.Get(elset.CurrentID())
			rc = nd.Get(el.Node(0))
			v(0) = nd.x
			rc = nd.Get(el.Node(1))
			v(3) = nd.x
			If max < v(0) Then max = v(0)
			If max < v(3) Then max = v(3)
			If min > v(0) Then min = v(0)
			If min > v(3) Then min = v(3)
			rc = el.Length(L)
			Length = Length + L
		Loop

		Dim val1 As String
		Dim val2 As String

		Do
			val1 = InputBox("Input the leftmost value of Load:", "Linear Distributed load")
			If val1 = "" Then GoTo Cleanup
			If IsNumeric(val1) Then Exit Do
		Loop

		Do
			val2 = InputBox("Input the rightmost value of Load:", "Linear Distributed load")
			If val2 = "" Then GoTo Cleanup
			If IsNumeric(val1) Then Exit Do
		Loop


		Dim eq As String, eq2 As String
		eq = CDbl(val1) & "+" & (CDbl(val2) - CDbl(val1)) & "*((X1+X0)/2-" & min & ")/" & max - min
		'eq = UCase(InputBox("Input the equation:" & vbLf & "variables: [x0,y0,z0] [x1,y1,z1] [A0,A1] [L]", "Load by equation"))
		'If eq = "" Then
		'	er = 2
		'	GoTo Cleanup
		'End If

		DOF(0) = GetDirection()
		If DOF(0) = 0 Then
			er = 3
			GoTo Cleanup
		End If

		rc = elset.Reset()
		Do While elset.Next() <> 0
			rc = el.Get(elset.CurrentID())
			rc = nd.Get(el.Node(0))
			v(0) = nd.x
			v(1) = nd.y
			v(2) = nd.z
			rc = nd.Get(el.Node(1))
			v(3) = nd.x
			v(4) = nd.y
			v(5) = nd.z
			eq2 = eq
			rc = el.Area(A0, A1)
			v(6) = CDbl(A0)
			v(7) = CDbl(A1)
			rc = el.Length(L)
			v(8) = CDbl(L)
			For i = 0 To UBound(v) Step 1
				If InStr(1, eq2, vecn(i)) <> 0 Then
					eq2 = Replace(eq2, vecn(i), CStr(v(i)), 1,-1)
				End If
			Next
			eq2 = Replace(eq2, ",", ".")
			If DBUG Then
				Debug.Print eq & " -> " & eq2 & " = " & Eval(eq2)
			End If
			load = CDbl(Eval(eq2))
			vals(0) = load
			vals(1) = load
			rc = ld.Add(-el.ID, 41, 0, DOF, vals, func)
			'Call App.feAppMessage(FCM_NORMAL, Eval(eq2))
		Loop
		Call App.feViewRegenerate(0)
Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set ld = Nothing
	Set nd = Nothing
	Set el = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub

Private Function GetDirection() As Long
	Dim rc As Long
	ldir = 0

	Begin Dialog UserDialog 320,91,"Get Load Direction",.loaddir ' %GRID:10,7,1,1
		GroupBox 10,7,210,84,"Direction",.GroupBox1
		OptionGroup .Group1
			OptionButton 20,28,90,14,"Element X",.Ex
			OptionButton 20,49,90,14,"Element Y",.Ey
			OptionButton 20,70,90,14,"Element Z",.Ez
			OptionButton 130,28,90,14,"Global X",.Gx
			OptionButton 130,49,90,14,"Global Y",.Gy
			OptionButton 130,70,90,14,"Global Z",.Gz
		OKButton 230,14,90,21
		CancelButton 230,42,90,21
	End Dialog
	Dim dlg As UserDialog
	dlg.Group1 = 5
	rc = Dialog(dlg)

	GetDirection = ldir
End Function

Rem See DialogFunc help topic for more information.
Private Function loaddir(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem loaddir = True ' Prevent button press from closing the dialog box
		loaddir = True
		Select Case DlgItem$
		Case "OK"
			ldir = DlgValue("Group1") + 1
			loaddir = False
		Case "Cancel"
			ldir = 0
			loaddir = False
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : loaddir = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
