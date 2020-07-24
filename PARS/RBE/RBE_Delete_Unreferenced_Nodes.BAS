Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim el As Object
    Set el = App.feElem()

    Dim elset As Object
    Set elset = App.feSet()

    Dim nd As Object
    Set nd = App.feNode()

    Dim ndcount As Long
    ndcount = 0
    Dim ndCountPerEl As Long


    Dim ndset As Object
    Set ndset = App.feSet()

    rc = elset.Select(FT_ELEM, True, "Select RBE3 Elements for Free Node Reduction")
    If rc <> -1 Then
    	er = 1
    	GoTo Cleanup
    End If

    Dim nodecount1 As Long, nodecount2 As Long
    Dim nds1, nds2
    Dim face1, face2
    Dim weight1, weight2
    Dim dof1, dof2

    rc = elset.Reset()
    Do While elset.Next() <> 0
    	rc = el.Get(elset.CurrentID())
		nodecount2 = 0
		ndCountPerEl = 0
    	If (el.type = 29) And (el.topology = 13) Then
			rc = el.GetNodeList(0, nodecount1, nds1, face1, weight1, dof1)
			For i = 0 To nodecount1 - 1 Step 1
				rc = ndset.Clear()
				rc = ndset.AddRule(nds1(i), FGD_ELEM_BYNODE)
				If ndset.Count() > 1 Then
					nodecount2 = nodecount2 + 1
					ReDim Preserve nds2(nodecount2 - 1)
					ReDim Preserve face2(nodecount2 - 1)
					ReDim Preserve weight2(nodecount2 - 1)
					ReDim Preserve dof2(nodecount2 * 6 - 1)

					nds2(nodecount2 - 1) = nds1(i)
					face2(nodecount2 - 1) = face1(i)
					weight2(nodecount2 - 1) = weight1(i)
					For j = 0 To 5 Step 1
						dof2((nodecount2 - 1) * 6 + j) = dof1(i * 6 + j)
					Next
				Else
					ndcount = ndcount + 1
					ndCountPerEl = ndCountPerEl + 1
				End If
			Next
			rc = el.PutNodeList(0, nodecount2, nds2, face2, weight2, dof2)
			rc = el.Put(el.ID)
			Call App.feAppMessage(FCM_NORMAL, "Element: " & el.ID & vbTab & ndCountPerEl & " Nodes Unreferenced.")
		End If
    Loop

    Call App.feAppMessage(FCM_NORMAL, ndcount & " Nodes Unreferenced in Total.")

	er = 0
Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
		Set ndset = Nothing
		Set el = Nothing
		Set elset = Nothing
		Set nd = Nothing
		Set App = Nothing
End Sub
