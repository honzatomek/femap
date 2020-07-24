Rem File: MoveNodesByDeform.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim fNode As Object
	Dim xData As Object
	Dim yData As Object
	Dim zData As Object

	Dim outSetID As Long
	Dim nodeID As Long

	Dim outset  As Object
	Set outset = App.feOutputSet

	'Get the current view
	Dim View As Object 'The FEMAP view object
	Dim viewID As Long
	Set View = App.feView
	rc = App.feAppGetActiveView(viewID)
	rc = View.Get(viewID)

	'Create and Element Object
	Set fNode = App.feNode()

	' Find out if there is an active output set and if so use it
	If (View.OutputSet > 0) Then
    	OutputSetID = View.OutputSet
    	rc = outset.Get(OutputSetID)

	    'Load the Output Vector in question
	    Set xData = outset.vector(2)
	    Set yData = outset.vector(3)
	    Set zData = outset.vector(4)

	    rc = fNode.First()

        While (FE_OK =  rc)
            'If the Element is a Bar (2) then put it's data in the spreadsheet

            fNode.x = fNode.x + xData.value(fNode.ID)
            fNode.y = fNode.y + yData.value(fNode.ID)
            fNode.z = fNode.z + zData.value(fNode.ID)

            rc = fNode.Put(fNode.ID)
            rc = fNode.Next()
        Wend
    	App.feViewRegenerate(0)
	End If
End Sub


