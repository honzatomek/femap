Attribute VB_Name = "Module2"
Option Explicit

Const DBUG As Boolean = False
Dim tr(5) As Boolean

Sub Main
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    On Error GoTo ExitHere

    Dim elset As Object
    Set elset = App.feSet()
    Dim el As Object
    Set el = App.feElem()
    Dim elID As Long
    Dim tmp As Variant
    Dim thexp As Double

    Dim nodecount As Long
    Dim n As Variant
	Dim f As Variant
	Dim w As Variant
	Dim DOF As Variant

	For i = 0 To 5 Step 1
		tr(i) = True
	Next

	Begin Dialog UserDialog 310,105,"RBE2 Nodes Releases Update",.rbe2dlg ' %GRID:10,7,1,1
		GroupBox 10,7,190,91,"Select DOF:",.GroupBox1
		CheckBox 20,28,90,14,"T1",.t1
		CheckBox 20,49,90,14,"T2",.t2
		CheckBox 20,70,90,14,"T3",.t3
		CheckBox 120,28,90,14,"R1",.r1
		CheckBox 120,49,90,14,"R2",.r2
		CheckBox 120,70,90,14,"R3",.r3
		OKButton 210,14,90,21
		CancelButton 210,42,90,21
	End Dialog
	Dim dlg As UserDialog
	Do While Dialog(dlg) = -1
		rc = elset.Select(8, True, "Select Rigid elements for DOF Update (RBE2)")
		If rc <> -1 Then GoTo ExitHere
		rc = elset.Reset()

		elID = elset.Next()

		Do While elID <> 0
			rc = el.Get(elID)
			If (el.type = 29) And (el.topology = 13) And (Not el.RigidInterpolate) Then ' (Not el.RigidInterpolate)
				rc = el.GetNodeList(0, nodecount, n, f, w, DOF)
				If DBUG Then
					Debug.Print "Element: " & el.ID
					For i = 0 To 5
						Debug.Print "Release " & i & ": " & el.release(0,i)
					Next
					For i = 0 To nodecount - 1 Step 1
						Debug.Print "Slave Node: " & n(i) & " Releases :" & DOF(6*i) & DOF(6*i + 1)& DOF(6*i + 2) & DOF(6*i +3) & DOF(6*i + 4) & DOF(6*i + 5)
					Next
				End If

				For i = 0 To 5
					el.release(0,i) = tr(i)
				Next
				For i = 0 To nodecount - 1 Step 1
					DOF(6*i) = tr(0)
					DOF(6*i + 1) = tr(1)
					DOF(6*i + 2) = tr(2)
					DOF(6*i +3) = tr(3)
					DOF(6*i + 4) = tr(4)
					DOF(6*i + 5) = tr(5)
				Next

				rc = el.PutNodeList(0, nodecount, n, f, w, DOF)

				rc = el.Put(el.ID)
			End If

			elID = elset.Next()
		Loop
	feFemap.feViewRegenerate(1)
	Loop

ExitHere:
	On Error Resume Next
	Set el = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function rbe2dlg(DlgItem$, Action%, SuppValue?) As Boolean
	Dim i As Long
	Select Case Action%
	Case 1 ' Dialog box initialization
		For i = 0 To 5 Step 1
			DlgValue(i+1, tr(i))
		Next
	Case 2 ' Value changing or button pressed
		Rem rbe2dlg = True ' Prevent button press from closing the dialog box
		rbe2dlg = True
		Select Case DlgItem$
		Case "OK"
			rbe2dlg = False
		Case "Cancel"
			rbe2dlg = False
		Case Else
			i = DlgNumber(DlgItem$)
			tr(i - 1) = DlgValue(i)
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
		'Debug.Print DlgItem$ & ", " & Action% & ", " & SuppValue?
	Case 5 ' Idle
		Rem Wait .1 : rbe2dlg = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
