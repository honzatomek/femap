Rem file: MultiMassElements.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feNode As femap.Node
	Dim nodeSet As femap.Set
	Dim propID As Long

	Dim feProp As femap.Prop
	Dim propSet As femap.Set
	Dim feElem As femap.Elem

	Set nodeSet = App.feSet
	Set propSet = App.feSet

	Set feElem = App.feElem
	Set feProp = App.feProp

	rc = nodeSet.Select(FT_NODE,True,"Select Nodes for Mass Elements")

	If rc = -1 Then 'return code FE_OK
	    'Lets see how many nodes were selected
    	nodeCount = nodeSet.count
	    If nodeCount > 0 Then
	        ' Let's get the property
    	    rc = propSet.SelectID( FT_PROP, "Select Mass Property", propID)
			If rc = 4 Then
				App.feAppMessageBox (0, "No Properties Exist, you will need to make a Mass Property and try again.")
				GoTo No_Mass
			End If
	        If rc = -1 Then
	            rc = feProp.Get(propID)
	            If rc = -1 Then
	                If feProp.type = FET_L_MASS	Then 'Mass
	                    rc = nodeSet.Reset
	                    nodeID = nodeSet.Next
	                    feElem.type = FET_L_MASS
    	                feElem.propID = propID
        	            feElem.topology = FTO_POINT1
	                    Do While nodeID <> 0
    	                    feElem.Node(0) = nodeID
        	                feElem.Put (feElem.NextEmptyID)
            	            nodeID = nodeSet.Next
                	    Loop
                	    Else
							App.feAppMessageBox (0, "The selected Property can not be used to create Mass elements.  Please select a Mass Property from the list or create one and try again.")
                	    	GoTo No_Mass
                	End If
            	End If
        	End If
    	End If
    	Call App.feViewRegenerate(0)
	End If

	No_Mass:
	
End Sub
