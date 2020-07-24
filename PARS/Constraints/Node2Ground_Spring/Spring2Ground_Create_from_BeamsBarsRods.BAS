Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim nset As Object
    Dim eset As Object
    Dim eset_s2g As Object
    Dim n As Object
    Dim els_by_node As Object
    Dim rc As Long, i As Long

   	Dim nID As Long
    Dim pID As Long
    Dim csID As Long
    Dim length As Double
    Dim pname As String
    Dim stiff_coeff As Variant

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Set eset = App.feSet()
    Set eset_s2g = App.feSet()
    Set nset = App.feSet()
    Set n = App.feNode()
    Set els_by_node = App.feSet()

    rc = eset.Select(FT_ELEM, True, "Select elements to be supported by Spring2Ground:")
    If (rc <> -1) Or (eset.Count() < 1) Then
		Call App.feAppMessage(FCM_NORMAL, "No Element selected.")
    	GoTo Cleanup
    End If

     'stiff_coeff = Array(0.38, 0.62, 1.00, 0.00, 0.00, 0.00)
    rc = get_stiffness_coefficients(App, stiff_coeff)
    If rc = FE_CANCEL Then GoTo Cleanup

    rc = eset_s2g.AddRule(1, FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(2, FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(5, FGD_ELEM_BYTYPE)
    rc = eset.RemoveNotCommon(eset_s2g.ID)

    nset.Clear()
	rc = nset.AddSetRule(eset.ID, FGD_NODE_ONELEM)
	nset.Reset()
	nID = nset.Next()
	Do While nID > 0
		rc = n.Get(nID)
		If rc = FE_FAIL Then
			Call App.feAppMessage(FCM_WARNING, "Node " & nID & " not found.")
			GoTo Cleanup
		End If

' 1 - Create a Set of elements on current Node, get their length and Property name
		rc = get_length_name(App, els_by_node, pname, length, nID, eset.ID)
		If rc = -1 Then
			er = 1
			GoTo Cleanup
		End If

' 2 - Create Coordination System from Line Elements
		rc = coordination_system_from_line_elements(App, els_by_node.ID, n.ID, pname)
		If rc = -1 Then
			er = 2
			GoTo Cleanup
		Else
			csID = rc
		End If

		n.outCSys =csID
		rc = n.Put(n.ID)

' 3 - Create Spring2Ground Property
		rc = spring_to_ground_property(App, pname, length, stiff_coeff, csID)
		If rc = -1 Then
			er = 3
			GoTo Cleanup
		Else
			pID = rc
		End If

' 4 - Create Spring2Ground element
		rc = spring_to_ground_element(App, n.ID, pID, csID)
		If rc = -1 Then
			er = 4
			GoTo Cleanup
		End If

		nID = nset.Next()
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next

	Set els_by_node = Nothing
	Set n = Nothing
	Set eset_s2g = Nothing
	Set eset = Nothing
	Set nset = Nothing
	Set App = Nothing
End Sub

Function get_stiffness_coefficients(ByRef App As femap.model, ByRef stiffness_coefficients As Variant) As Long
	Dim rc As Long
	Dim vecBase1 As Variant, vecBase2 As Variant

	On Error GoTo Cleanup

	App.Info_DefaultX = 1.0
	App.Info_DefaultY = 1.0
	App.Info_DefaultZ = 1.0
	rc = App.feCoordPickByMethod(FCD_COORDINATES, "Input Translational Stiffness Coefficients:", vecBase1)
	If rc = FE_CANCEL Then GoTo Cleanup

	App.Info_DefaultX = 1.0
	App.Info_DefaultY = 1.0
	App.Info_DefaultZ = 1.0
	rc = App.feCoordPickByMethod(FCD_COORDINATES, "Input Rotational Stiffness Coefficients:", vecBase2)
	If rc = FE_CANCEL Then GoTo Cleanup

	stiffness_coefficients = Array(vecBase1(0), vecBase1(1), vecBase1(2), vecBase2(0), vecBase2(1), vecBase2(2))

Cleanup:
	get_stiffness_coefficients = rc
End Function

Function get_length_name(ByRef App As femap.model, ByRef eset As Object, ByRef ele_name As String, ByRef length As Double, ByVal nodeID As Long, ByVal base_element_set_ID As Long) As Long
	Dim rc As Long
	Dim e As Object
	Dim p As Object
	Dim eID As Long
	Dim tmp As Double

	On Error GoTo Cleanup

	rc = -1
	Set eset = App.feSet()
	Set e = App.feElem()
	Set p = App.feProp()
	ele_name = "S2G"
	length = 0.0

	eset.Clear()
	rc = eset.AddRule(nodeID, FGD_ELEM_BYNODE)
	If (rc = FE_FAIL) Or (eset.Count() < 1) Then
		rc = -1
		Call App.feAppMessage(FCM_NORMAL, "No Elements on Node " & nodeID & ".")
		GoTo Cleanup
	End If
	rc = eset.RemoveNotCommon(base_element_set_ID)
	If rc = FE_FAIL Then
		rc = -1
		GoTo Cleanup
	End If

	length = 0.0
	eset.Reset()
	eID = eset.Next()
	Do While eID > 0
		rc = e.Get(eID)
		If rc = FE_FAIL Then
			rc = -1
			Call App.feAppMessage(FCM_NORMAL, "Element " & eID & " not found.")
			GoTo Cleanup
		End If
		If e.type = 1 Or e.type = 2 Or e.type = 5 Then
			rc = e.Length(tmp)
			If rc = FE_FAIL Then
				rc = -1
				GoTo Cleanup
			End If
			length = length + tmp

			rc = p.Get(e.propID)
			If rc = FE_FAIL Then
				rc = -1
				GoTo Cleanup
			End If
			If p.title <> ele_name Then
				ele_name = ele_name & "_" & p.title
			End If
		End If
		eID = eset.Next()
	Loop
	length = length / 2
	rc = eset.ID

Cleanup:
	On Error Resume Next
	If rc = -1 Then
		Call App.feAppMessage(FCM_WARNING, "Problem With Line Elements on Node " & nodeID & ".")
	End If
	Set e = Nothing
	Set p = Nothing
	get_length_name = rc
End Function


Function coordination_system_from_line_elements(ByRef App As femap.model, ByVal element_set_ID As Long, ByVal nodeID As Long, ByVal csys_name As String) As Long
	''' Creates a CSys from Line element end_points, returns new CSys ID on success, -1 on failure
	Dim rc As Long
	Dim nset As Object
	Dim n As Object
	Dim cs As Object
	Dim csID
	Dim vx As Variant, vy As Variant ', vz As Variant
	Dim nID As Long

	On Error GoTo Cleanup
	rc = -1

	Set nset = App.feSet()
	Set n = App.feNode()
	Set cs = App.feCSys()

	csID = cs.NextEmptyID()

	nset.Clear()
	rc = nset.AddSetRule(element_set_ID, FGD_NODE_ONELEM)
	Select Case nset.Count()
	Case 0, 1
		Call App.feAppMessage(FCM_NORMAL, csys_name & " CSys creation: element set has " & nset.Count() & " nodes.")
		rc = -1
		GoTo Cleanup
	Case 2
	Case 3
		rc = nset.Remove(nodeID)
	Case Else
		Call App.feAppMessage(FCM_NORMAL, csys_name & " CSys creation: element set has " & nset.Count() & " nodes, only first two will be used.")
		rc = nset.Remove(nodeID)
	End Select

	nset.Reset()
	nID = nset.Next()
	rc = n.Get(nID)
	vx =Array(n.x, n.y, n.z)

	nID = nset.Next()
	rc = n.Get(nID)
	vx = Array(n.x - vx(0), n.y - vx(1), n.z - vx(2))

	rc = App.feVectorCrossProduct(Array(0.0,0.0,1.0), vx, vy)
	rc =App.feVectorUnit(vx, vx)
	rc =App.feVectorUnit(vy, vy)

	rc = n.Get(nodeID)
	rc = cs.Get(0)
	rc = cs.XYPoints(n.xyz, Array(n.x + vx(0), n.y + vx(1), n.z + vx(2)), Array(n.x + vy(0), n.y + vy(1), n.z + vy(2)))
	cs.title = csys_name

	rc = cs.Put(csID)
	If rc = FE_FAIL Then
		rc = -1
	Else
		rc = cs.ID
	End If


Cleanup:
	On Error Resume Next
	If rc = -1 Then
		Call App.feAppMessage(FCM_WARNING, "CSys from Line Elements " & csys_name & " failed.")
	End If
	Set nset = Nothing
	Set n = Nothing
	Set cs = Nothing
	coordination_system_from_line_elements = rc
End Function


Function spring_to_ground_property(ByRef App As femap.model, ByVal property_name As String,  _
														         ByVal base_stiffness As Double, ByVal stiffness_coefficients As Variant, _
														         ByVal csysID As Long) As Long
	''' Creates a Spring To Ground Property, returns new Property ID on success, -1 on failure
	Dim rc As Long, i As Long
	Dim p As Object
	Dim pID As Long

	On Error GoTo Cleanup
	rc = -1

	Set p = App.feProp()

	pID = p.NextEmptyID()

	p.title = property_name
	p.type = 41
	p.cbush = 1
	p.flag(0) = csysID
	p.refCS = csysID
	For i = 0 To 5 Step 1
		p.pval(36 + i) = base_stiffness * stiffness_coefficients(i)
	Next
	rc = p.Put(pID)
	If rc = FE_FAIL Then
		rc = -1
	Else
		rc = p.ID
	End If

Cleanup:
	On Error Resume Next
	If rc = -1 Then
		Call App.feAppMessage(FCM_WARNING, "Spring2Ground Property " & property_name & " failed.")
	End If
	Set p = Nothing
	spring_to_ground_property = rc
End Function


Function spring_to_ground_element(ByRef App As femap.model, ByVal nodeID As Long, ByVal propID As Long, ByVal csysID As Long) As Long
	''' Creates a Spring To Ground Element, returns new Element ID on success, -1 on failure
	Dim rc As Long
	Dim e As Object

	On Error GoTo Cleanup
	rc = -1

	Set e = App.feElem()

	e.type = 41
	e.propID = propID
	e.topology = 9
	e.Node(0) = nodeID
	e.SpringUseCID = True
	e.SpringCID = csysID

	rc = e.Put(e.NextEmptyID())

	If rc = FE_FAIL Then
		rc = -1
	Else
		rc = e.ID
	End If

Cleanup:
	On Error Resume Next
	If rc = -1 Then
		Call App.feAppMessage(FCM_WARNING, "Spring2Ground Element at Node " & nodeID & " failed.")
	End If
	Set e = Nothing
	spring_to_ground_element = rc
End Function
