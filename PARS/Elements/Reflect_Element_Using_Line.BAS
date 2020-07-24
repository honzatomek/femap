Option Explicit

'Reflects an element using a line perpendicular to reflection plane
'where the origin of the plane Is the Line midpoint

Private er As Long
Const DBUG As Boolean = True

Sub Main()
	Dim App As femap.model
	Set App = feFemap()
	Dim rc As Long, i As Long

	On Error GoTo Cleanup
	er = 0

	Dim elset As Object
	Set elset = App.feSet()

	Dim ndset As Object
	Set ndset = App.feSet()

	Dim ndID As Long

	Dim nd1 As Object
	Set nd1 = App.feNode()

	Dim nd2 As Object
	Set nd2 = App.feNode()

	Dim midnd(2) As Double
	Dim norm(2) As Double

	rc = elset.Select(8, True, "Select elements to be reflected:")
	If (rc <> -1) Or (elset.Count < 1) Then
		er = 1
		GoTo Cleanup
	End If

	rc = ndset.SelectID(7, "Select first node to create plane.", ndID)
	If (rc<>-1) Or (ndset.Count() < 1) Then
		er = 2
		GoTo Cleanup
	End If
	rc = nd1.Get(ndID)

	rc = ndset.SelectID(7, "Select second node to create plane.", ndID)
	If (rc<>-1) Or (ndset.Count() < 2) Then
		er = 3
		GoTo Cleanup
	End If
	rc = nd2.Get(ndID)

	midnd(0) = (nd1.x() + nd2.x())/2
	midnd(1) = (nd1.y() + nd2.y())/2
	midnd(2) = (nd1.z() + nd2.z())/2

	norm(0) = nd2.x() - nd1.x()
	norm(1) = nd2.y() - nd1.y()
	norm(2) = nd2.z() - nd1.z()

	rc = App.feGenerateReflect(FT_ELEM, elset.ID, midnd, norm, 0, False)
	If rc <> -1 Then
		er = 4
		GoTo Cleanup
	End If

Cleanup:
	On Error Resume Next
		Set nd2 = Nothing
		Set nd1 = Nothing
		Set ndset = Nothing
		Set elset = Nothing
		If DBUG Then App.feAppMessage(FCM_NORMAL, "The script exited with code " & er)
		Set App = Nothing
End Sub
