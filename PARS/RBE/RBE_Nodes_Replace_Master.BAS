Option Explicit

'Script to replace rigid spider nodes
Const DBUG As Boolean = False
Private StartTime As Double

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    If DBUG Then
    	On Error GoTo 0
    Else
    	On Error GoTo ExitHere
    End If

    StartTime = Timer

    Dim el As Object
    Set el = App.feElem
    Dim elID As Long

    Dim s As Object
    Set s = App.feSet()
    Dim ns As Object
    Set ns = App.feSet()

    Dim ndold As Long
    Dim ndnew As Long

    Dim ndcount As Long
    Dim nds As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim DOF As Variant

    Dim c As Long
    c = 0

    Do While s.SelectID(8, "Select RBE2/RBE3 Element for nodes Update:", elID) = -1
		rc = ns.Clear()
    	rc = el.Get(elID)
		pt("Selected element: " & elID & ", type = " & el.type & ", topology = " & el.topology, c + 1)
		If (el.type = 29) And (el.topology = 13) Then
			c = c + 1
			'rc = el.GetNodeList(0, ndcount, nds, face, weight, DOF)
			ndold = el.Node(0)
			'Do
				'rc = ns.SelectID(7, "Select Node to be Replaced:", ndold)
				'If rc <> -1 Then GoTo ExitHere
				'For i = 0 To (ndcount - 1) Step 1
				'	If nds(i) = ndold Then Exit Do
				'Next
			'Loop

			rc = ns.SelectID(7, "Select new node:", ndnew)
			If rc <> -1 Then GoTo ExitHere
			'nds(i) = ndnew
			'rc = el.PutNodeList(0, ndcount, nds, face, weight, DOF)
			el.Node(0) = ndnew
			rc = el.Put(elID)
			If rc = -1 Then pt("Updated element " & elID & " by changing node " & ndold & " to node " & ndnew & ".", c)
			feFemap.feViewRegenerate(1)
		End If
    Loop

    If c > 0 Then Call feFemap.feViewShow(8, s.ID)

ExitHere:
	On Error Resume Next
		Set el = Nothing
		Set ns = Nothing
		Set s = Nothing
		Set App = Nothing
End Sub

Private Sub pt(ByVal message As String, Optional ByVal i As Long = 0, Optional ByVal debugging As Boolean = False)
	'Dim App As femap.model
	'Set App = feFemap
	Dim rc As Long

	If DBUG Then Debug.Print "["& i & "] - " & MinutesElapsed & " - " & message
	If (Not debugging) Then	rc = feFemap.feAppMessage(FCM_NORMAL, "["& i & "] - " & MinutesElapsed & " - " & message)

	'Set App = Nothing
End Sub

Private Function MinutesElapsed(Optional ByVal t As Double = 0) As String
	If t = 0 Then t = StartTime
	MinutesElapsed = Format((Timer - t) / 86400, "hh:mm:ss")
End Function
