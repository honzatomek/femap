Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Begin Dialog UserDialog 300,154 ' %GRID:10,7,1,1
		CheckBox 40,21,90,14,"Update X",.CheckBox1,1
		CheckBox 40,42,90,14,"Update Y",.CheckBox2,1
		CheckBox 40,63,90,14,"Update Z",.CheckBox3,1
		TextBox 140,77,120,21,.TextBox1
		text 30,84,90,14,"Scale Factor:",.Text1,1
		OKButton 40,112,90,21
		CancelButton 140,112,90,21
	End Dialog

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

   		Dim dlg As UserDialog

		dlg.CheckBox1 = 1
		dlg.CheckBox2 = 1
		dlg.CheckBox3 = 1
		dlg.TextBox1 = "1.00"

        If Dialog(dlg) = 0 Then
           GoTo JUMPING_OUT
        End If

        Dim doX As Boolean
        Dim doY As Boolean
        Dim doZ As Boolean
        Dim sFactor As Double

        If dlg.CheckBox1 = 1 Then
           doX = True
        Else
           doX = False
        End If

        If dlg.CheckBox2 = 1 Then
           doY = True
        Else
           doY = False
        End If

        If dlg.CheckBox3 = 1 Then
           doZ = True
        Else
           doZ = False
        End If

        sFactor = Val( dlg.TextBox1 )

        If dlg.CheckBox1 + dlg.CheckBox2 + dlg.CheckBox3 = 0 Then
           MsgBox( "No Components Selected, Exiting...", vbOkOnly, "Move Nodes by Deformation" )
           GoTo JUMPING_OUT
        End If

    	Msg = "For Output Set #" + Str$(OutputSetID)
    	If doX Then
    		Msg = Msg + " updating X-Value"
    	End If
    	If doY Then
    		Msg = Msg + " updating Y-Value"
    	End If
		If doZ Then
    		Msg = Msg + " updating Z-Value"
    	End If

    	Msg = Msg + " scaled by " + Str$( sFactor)
    	rc = App.feAppMessage(FCM_NORMAL, Msg )

	    'Load the Output Vector in question
	    Set xData = outset.Vector(2)
	    Set yData = outset.Vector(3)
	    Set zData = outset.Vector(4)

	    rc = fNode.First()

        While (FE_OK =  rc)
            'If the Element is a Bar (2) then put it's data in the spreadsheet
			If doX Then
	            fNode.x = fNode.x + sFactor * xData.Value(fNode.ID)
	        End If
            If doY Then
	            fNode.y = fNode.y + sFactor * yData.Value(fNode.ID)
	        End If
	        If doZ Then
            	fNode.z = fNode.z + sFactor * zData.Value(fNode.ID)
            End If
            rc = fNode.Put(fNode.ID)
            rc = fNode.Next()
        Wend
    	App.feViewRegenerate(0)
	Else
		MsgBox( "No Current Output Set in the Active View, Exiting.....", vbOkOnly, "Move Nodes by Deformation")
	End If
	JUMPING_OUT:

End Sub


