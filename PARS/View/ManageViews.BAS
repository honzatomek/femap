Option Explicit
'Creates Views in +XY, -XY, +XZ, -XZ, +YZ, -YZ and axonometry

Const DBUG As Boolean = True
Private rotate(6) As Variant
Private names(6) As String
Private curview As Long

Sub Main()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long

	On Error GoTo ExitHere

	rotate(0) = Array(0,0,0) '+XY
	rotate(1) = Array(180,0,0) '-XY
	rotate(2) = Array(-90,0,0) '+XZ
	rotate(3) = Array(-90,0,180) '-XZ
	rotate(4) = Array(-90,0,-90) '+YZ
	rotate(5) = Array(-90,0,90) '-YZ
	rotate(6) = Array(-75,0,-135) 'axonometry

	names(0) = "+XY"
	names(1) = "-XY"
	names(2) = "+XZ"
	names(3) = "-XZ"
	names(4) = "+YZ"
	names(5) = "-YZ"
	names(6) = "axonometry"

	curview = App.Info_ActiveID(FT_VIEW)

	Begin Dialog UserDialog 340,147,"Views Manager",.viewmanfunc ' %GRID:10,7,1,1
		GroupBox 10,7,210,133,"Select View:",.GroupBox1
		PushButton 20,28,90,21,"+XY",.PushButton1
		PushButton 120,28,90,21,"-XY",.PushButton2
		PushButton 20,56,90,21,"+XZ",.PushButton3
		PushButton 120,56,90,21,"-XZ",.PushButton4
		PushButton 20,84,90,21,"+YZ",.PushButton5
		PushButton 120,84,90,21,"-YZ",.PushButton6
		PushButton 20,112,190,21,"Axonometry",.PushButton7
		CancelButton 230,112,100,21
		PushButton 230,28,100,21,"Create Views",.create
		PushButton 230,56,100,21,"Reset Views",.rset
	End Dialog
	Dim dlg As UserDialog
	rc = Dialog(dlg)

ExitHere:
	On Error Resume Next
		Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function viewmanfunc(DlgItem$, Action%, SuppValue?) As Boolean
	Dim i As Long
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem viewmanfunc = True ' Prevent button press from closing the dialog box
		viewmanfunc = True
		Select Case DlgItem$
		Case "Cancel"
			viewmanfunc = False
		Case "create"
			Call CreateViews()
			viewmanfunc = False
		Case "rset"
			Call ReSetViews()
			viewmanfunc = False
		Case Else
			i = DlgNumber(DlgItem$) - 1
			If (i>-1) And (i<7) Then
				Call RotateCurrent(rotate(i)(0), rotate(i)(1), rotate(i)(2))
				viewmanfunc = False
			End If
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : viewmanfunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


Private Sub RotateCurrent(ByVal rx As Double, ByVal ry As Double, ByVal rz As Double)
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long, j As Long

	On Error GoTo ExitHere

	Dim v As Object
	Set v = App.feView()

	rc = v.Get(curview)
	If v.ID = -1 Then rc = v.Get(1)

	rc = v.AlignCSys(0)
	rc = v.Rotate(0, 0, 0, 1, 0, 0, rx)
	rc = v.Rotate(0, 0, 0, 0, 1, 0, ry)
	rc = v.Rotate(0, 0, 0, 0, 0, 1, rz)
	rc = v.AutoscaleAll(True)
	rc = v.Put(v.ID())
	rc = v.Redraw()

ExitHere:
	On Error Resume Next
		Set v = Nothing
		Set App = Nothing
End Sub

Private Sub CreateViews()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long, j As Long

	On Error GoTo ExitHere

	Dim rotate(6) As Variant
	Dim names(6) As String
	Dim exists As Boolean

	rotate(0) = Array(0,0,0) '+XY
	rotate(1) = Array(180,0,0) '-XY
	rotate(2) = Array(-90,0,0) '+XZ
	rotate(3) = Array(-90,0,180) '-XZ
	rotate(4) = Array(-90,0,-90) '+YZ
	rotate(5) = Array(-90,0,90) '-YZ
	rotate(6) = Array(-75,0,-135) 'axonometry

	names(0) = "+XY"
	names(1) = "-XY"
	names(2) = "+XZ"
	names(3) = "-XZ"
	names(4) = "+YZ"
	names(5) = "-YZ"
	names(6) = "axonometry"

	Dim vorg As Object
	Set vorg = App.feView()

	Dim v As Object
	Set v = App.feView()

	rc = vorg.Get(1)
	If DBUG Then Debug.Print vorg.ID()
	For i = 0 To 6 Step 1
		exists = False
		For j = 1 To vorg.CountSet()
			rc = v.Get(j)
			If names(i) = v.title Then exists = True
		Next
		If Not exists Then
			rc = v.Get(1)
			v.title = names(i)
			rc = v.AlignCSys(0)
			rc = v.Rotate(0, 0, 0, 1, 0, 0, rotate(i)(0))
			rc = v.Rotate(0, 0, 0, 0, 1, 0, rotate(i)(1))
			rc = v.Rotate(0, 0, 0, 0, 0, 1, rotate(i)(2))
			rc = v.AutoscaleAll(True)
			rc = v.Put(v.NextEmptyID())
			rc = v.Redraw()
		End If
	Next

ExitHere:
	On Error Resume Next
		Set v = Nothing
		Set vorg = Nothing
		Set App = Nothing
End Sub

Private Sub ReSetViews()
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long, j As Long

	On Error GoTo ExitHere

	Dim vorg As Object
	Set vorg = App.feView()

	Dim v As Object
	Set v = App.feView()

	rc = vorg.Get(1)
	If DBUG Then Debug.Print vorg.ID()
	For i = 0 To 6 Step 1
		For j = 1 To vorg.CountSet()
			rc = v.Get(j)
			If names(i) = v.title Then
				rc = v.AlignCSys(0)
				rc = v.Rotate(0, 0, 0, 1, 0, 0, rotate(i)(0))
				rc = v.Rotate(0, 0, 0, 0, 1, 0, rotate(i)(1))
				rc = v.Rotate(0, 0, 0, 0, 0, 1, rotate(i)(2))
				rc = v.AutoscaleAll(True)
				rc = v.Put(v.ID())
				rc = v.Redraw()
				Exit For
			End If
		Next
	Next

ExitHere:
	On Error Resume Next
		Set v = Nothing
		Set vorg = Nothing
		Set App = Nothing
End Sub
