Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    'Dim etest As Object
    'Set etest =App.feElem()
    'rc =etest.Get(904)

    Dim eset As Object
    Set eset = App.feSet()
    Dim eset_s2g As Object
    Set eset_s2g = App.feSet()

    Dim nset As Object
    Set nset = App.feSet()
    Dim nID As Long, nID2 As Long
    Dim n As Object
    Set n = App.feNode()
    Dim eID As Long
    Dim e As Object
    Set e = App.feElem()

    Dim els_by_node As Object
    Set els_by_node = App.feSet()

    Dim nds_by_elems As Object
    Set nds_by_elems = App.feSet()

    Dim p As Object
    Set p = App.feProp()
    Dim pID As Long
    pID = 10000

    Dim cs As Object
    Set cs = App.feCSys()
    Dim csID As Long
    csID = 10000

    Dim length As Double
    Dim pname As String
    Dim tmp As Double
    Dim n2 As Object
    Set n2 = App.feNode()
    Dim vx As Variant, vy As Variant, vz As Variant

    Dim stiff_coeff As Variant
    stiff_coeff = Array(0.38, 0.62, 1.00, 0.00, 0.00, 0.00)

    rc = eset.Select(FT_ELEM, True, "Select elements to be supported by Spring2Ground:")
    If (rc <> -1) Or (eset.Count() < 1) Then
		Call App.feAppMessage(FCM_NORMAL, "No Element selected.")
    	GoTo Cleanup
    End If

    rc = eset_s2g.AddRule(1, FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(2, FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(5, FGD_ELEM_BYTYPE)
    rc = eset.RemoveNotCommon(eset_s2g.ID)

    nset.Clear()
	rc = nset.AddSetRule(eset.ID, FGD_NODE_ONELEM)
	nset.Reset()
	nID = nset.Next()
	Do While nID > 0
		Debug.Print nID
		rc = n.Get(nID)
		If rc = FE_FAIL Then
			Call App.feAppMessage(FCM_WARNING, "Node " & nID & " not found.")
			GoTo Cleanup
		End If

		els_by_node.Clear()
		rc = els_by_node.AddRule(nID, FGD_ELEM_BYNODE)
		If (rc = FE_FAIL) Or (els_by_node.Count() < 1) Then
			Call App.feAppMessage(FCM_WARNING, "No Elements on Node " & nID & ".")
			GoTo Cleanup
		End If
		rc = els_by_node.RemoveNotCommon(eset.ID)

		length = 0.0
		pname = "S2G"
		els_by_node.Reset()
		eID = els_by_node.Next()
		Do While eID > 0
			rc = e.Get(eID)
			If rc = FE_FAIL Then
				Call App.feAppMessage(FCM_WARNING, "Element " & eID & " not found.")
				GoTo Cleanup
			End If
			If e.type = 1 Or e.type = 2 Or e.type = 5 Then
				rc = e.Length(tmp)
				length = length + tmp

				rc = p.Get(e.propID)
				If p.title <> pname Then
					pname = pname & "_" & p.title
				End If
			End If
			eID = els_by_node.Next()
		Loop
		length = length / 2

		nds_by_elems.Clear()
		rc = nds_by_elems.AddSetRule(els_by_node.ID ,FGD_NODE_ONELEM)
		If nds_by_elems.Count() <> 2 Then
			rc = nds_by_elems.Remove(nID)
		End If

		nds_by_elems.Reset()
		nID2 = nds_by_elems.Next()
		rc = n2.Get(nID2)
		vx =Array(n2.x, n2.y, n2.z)

		nID2 = nds_by_elems.Next()
		rc = n2.Get(nID2)
		vx(0) = n2.x - vx(0)
		vx(1) = n2.y - vx(1)
		vx(2) = n2.z - vx(2)

		rc = App.feVectorCrossProduct(Array(0.0,0.0,1.0), vx, vy)
		rc = App.feVectorCrossProduct(vx, vy, vz)
		rc =App.feVectorUnit(vx, vx)
		rc =App.feVectorUnit(vy, vy)
		rc =App.feVectorUnit(vz, vz)

		rc = cs.Get(0)
		rc = cs.XYPoints(n.xyz, Array(n.x + vx(0), n.y + vx(1), n.z + vx(2)), Array(n.x + vy(0), n.y + vy(1), n.z + vy(2)))
		cs.title = pname
		If cs.NextEmptyID() < csID Then
			csID = csID + 1
		Else
			csID = cs.NextEmptyID()
		End If
		rc = cs.Put(csID)

		n.outCSys =cs.ID
		rc = n.Put(n.ID)

		Set p = App.feProp()
		p.title = pname
		p.type = 41
		p.cbush = 1
		p.flag(0) = cs.ID
		p.refCS = cs.ID
		For i = 0 To 5 Step 1
			p.pval(36 + i) = length * stiff_coeff(i)
		Next
		If p.NextEmptyID() < pID Then
			pID = pID + 1
		Else
			pID = p.NextEmptyID()
		End If
		rc = p.Put(pID)

		Set e = App.feElem()
		e.type = 41
		e.propID = p.ID
		e.topology = 9
		e.Node(0) = n.ID
		e.SpringUseCID = True
		e.SpringCID = cs.ID
		rc = e.Put(e.NextEmptyID())

		nID = nset.Next()
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub
