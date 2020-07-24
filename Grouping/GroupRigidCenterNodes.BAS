Sub CreateRigidNodeGroup ( App, feGroup, grID )
	feGroup.Put ( grID )
    feGroup.Get ( grID )
   	feGroup.title = "RBE Element Independent Nodes"
End Sub

Sub Main
	Dim App As femap.model
	Set App = GetObject(,"femap.model")

	Dim feElem As femap.Elem
    Set feElem = App.feElem
    Dim eID As Long
    Dim nID As Long

    Dim RBEs As femap.Set
    Set RBEs = App.feSet

    While feElem.Next
    	If feElem.type = FET_L_RIGID Then
			RBEs.Add ( feElem.ID )
        End If
    Wend

   	Dim feGroup As femap.Group
	Set feGroup = App.feGroup
    Dim grID As Long
    Dim activeGrID As Long
    activeGrID = App.feGroup.Active

    If RBEs.Count = 0 Then
    	GoTo EndMacro
    End If
    If activeGrID > 0 Then
    	rc = App.feAppMessageBox ( 2, "Add Nodes to Active Group?" )
		If rc = 0 Then
			grID = App.feGroup.NextEmptyID
			CreateRigidNodeGroup ( App, feGroup, grID )
		Else
			grID = activeGrID
			feGroup.Get ( grID )
		End If
	Else
		grID = App.feGroup.NextEmptyID
		CreateRigidNodeGroup ( App, feGroup, grID )
    End If

    RBEs.Reset
    eID = RBEs.First
    For i = 0 To RBEs.Count-1
   		feElem.Get ( eID )
   		nID = feElem.Nodes ( 0 )
		feGroup.Add ( FT_NODE, nID )
		feElem.Put ( eID )
		eID = RBEs.Next
    Next i
	feGroup.Put( grID )
   	Msg = Str$(RBEs.Count) + " Rigid Independent Nodes Placed in Group " + Str$(grID)
    rc = App.feAppMessage( FCM_NORMAL, Msg )
    rc = feGroup.Put( grID )

    EndMacro:
    App.feGroup.Active = activeGrID
	App.feSet.DeleteAll
    App.feViewRegenerate ( 0 )

End Sub
