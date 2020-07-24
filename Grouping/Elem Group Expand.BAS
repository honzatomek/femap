Rem File: ElemGroupExpand.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim groupID As Long
	Dim viewID As Long
	Dim elemSET As femap.Set
	Dim nodeSET As femap.Set
	Dim feGroup As femap.Group

	Set elemSET = App.feSet()
	Set nodeSET = App.feSet()
	Set feGroup = App.feGroup()

	groupID = 0

	groupID = App.Info_ActiveID(24)

	If groupID > 0 Then
    	rc = elemSET.AddGroup(8, groupID)
    	If rc = -1 Then
        	rc = nodeSET.AddSetRule(elemSET.ID, 20)
        	If rc = -1 Then
            	rc = elemSET.AddSetRule(nodeSET.ID, 25)
            	If rc = -1 Then
	                rc = nodeSET.AddSetRule(elemSET.ID, 20)
    	            If rc = -1 Then
        	            rc = feGroup.Get(groupID)
            	        If rc = -1 Then
                	        rc = feGroup.SetAdd(7, nodeSET.ID)
                    	    rc = feGroup.SetAdd(8, elemSET.ID)
                        	rc = feGroup.Put(groupID)
                        	viewID = App.Info_ActiveID(22)
                        	rc = App.feViewRegenerate(viewID)
                    	End If
                	End If
            	End If
        	End If
    	End If
	End If

	
End Sub
