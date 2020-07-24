Option Explicit On

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim rc As Long, i As Long

	'On Error GoTo Cleanup

	Dim eset As Object
	Set eset = App.feSet()

	Dim el As Object
	Set el = App.feElem()

	Dim area1 As Double
	Dim area2 As Double
	Dim area3 As Double
	Dim area4 As Double
	Dim area5 As Double
	Dim area6 As Double
	Dim areatotal As Double
	areatotal = 0

	rc = eset.Select(8, True, "Select planar elements for area inquiry:")
	If rc <> -1 Then GoTo Cleanup
	Call App.feAppMessage(FCM_NORMAL, eset.Count() & " elements selected..")

	rc = eset.Reset()
	For i = 0 To eset.Count() - 1 Step 1
		rc = el.Get(eset.Next())
		If rc <> -1 Then GoTo Cleanup
		If (el.type > 9) And (el.type < 23) Then
			el.GetFaceArea(1, area1)
			el.GetFaceArea(2, area2)
			el.GetFaceArea(3, area3)
			el.GetFaceArea(4, area4)
			el.GetFaceArea(5, area5)
			Select Case el.topology
			Case 4, 5
				el.GetFaceArea(6, area6)
			Case 2, 3
				area6 = 0
			End Select
			areatotal = areatotal + Max(Max(Max(area1, area2), Max(area3, area4)), Max(area5, area6))
		End If
	Next

	Call App.feAppMessage(FCM_NORMAL, "Accumulated Plate Element Area: " & areatotal)

Cleanup:
	On Error Resume Next
		Set el = Nothing
		Set eset = Nothing
		Set App = Nothing
    
End Sub

Private Function Max(ByVal A As Double, ByVal B As Double) As Double
	If A > B Then
		Max = A
	Else
		Max = B
	End If
End Function
