Rem File: MoveMergeNodes.BAS
Sub Main

	Dim App As femap.model
    Set App = feFemap()

	Begin Dialog UserDialog 320,119,"Move/Merge Nodes" ' %GRID:10,7,1,0
		CheckBox 20,18,20,28,"",.CheckBox1
		CheckBox 20,56,300,28,"  Merge Nodes",.CheckBox12
		PushButton 50,91,90,21,"Start",.PushButton1
		PushButton 180,91,90,21,"Cancel",.PushButton2
		text 50,35,230,12,"Move Both Nodes To Centroid (Off)",.Text1
		text 50,17,220,12,"Move Node 1 (On)",.Text12
	End Dialog
	Dim dlg As UserDialog
	rc = Dialog( dlg )

	If rc = 2 Then
		GoTo DONE
	End If

' Declare nodal data
'		nd1 is the first node - ID1 is the id of the first node
Dim nd1 As femap.Node
Dim ID1 As Long
'		nd2 is the second node - ID2 is the id of the second node
Dim nd2 As femap.Node
Dim ID2 As Long

Dim feSet As femap.Set
Dim newx As Double
Dim newy As Double
Dim newz As Double
Dim ndCount As Long
Dim mergedNodes As Variant

Set nd1 = App.feNode()
Set nd2 = App.feNode()
Set feSet = App.feSet()

DO_IT_AGAIN:

rc = feSet.SelectID(7, "Select Node 1", ID1)

If rc = 2 Then
    GoTo DONE
End If

GET_NODE_2:

rc = feSet.SelectID(7, "Select Node 2", ID2)

If rc = 2 Then
    GoTo DONE
End If

rc = nd1.Get(ID1)

If rc = 0 Then
    rc = MsgBox("Node 1 does not exist", vbOkOnly, "Error Selecting Node 1")
    GoTo DONE
End If

rc = nd2.Get(ID2)

If rc = 0 Then
    rc = MsgBox("Node 2 does not exist", vbOkOnly, "Error Selecting Node 2")
    GoTo DONE
End If

If ID1 = ID2 Then
    MsgStr$ = "Node " + Str$(ID1) + " selected twice, please try again."
    rc = MsgBox(MsgStr$, vbOkOnly, "Duplicate Selection")
    GoTo GET_NODE_2
End If

If dlg.CheckBox1 Then 'Move 1 to 2

	nd1.x = nd2.x
    nd1.y = nd2.y
    nd1.z = nd2.z

    rc = nd1.Put(ID1)
Else 'Average
    newx = (nd1.x + nd2.x) / 2#
    newy = (nd1.y + nd2.y) / 2#
    newz = (nd1.z + nd2.z) / 2#

    nd1.x = newx
    nd1.y = newy
    nd1.z = newz

    rc = nd1.Put(ID1)

    nd2.x = newx
    nd2.y = newy
    nd2.z = newz

    rc = nd2.Put(ID2)
End If


If dlg.CheckBox12 Then ' Merge

	    rc = App.feCheckCoincidentNode(-ID2, -ID1, 0.001, 1, 1, ndCount, mergedNodes)

End If

rc = App.feViewRegenerate(0)

GoTo DO_IT_AGAIN

DONE:

End Sub

