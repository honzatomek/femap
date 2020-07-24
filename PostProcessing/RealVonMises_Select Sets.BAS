Sub Main
	Dim App As femap.model
	Set App = GetObject(,"femap.model")

    Dim ouSet As femap.OutputSet
    Set ouSet = App.feOutputSet

    Dim oset As femap.Set
    Set oset = App.feSet

    oset.Select (FT_OUT_CASE, True, "Select Output To create Von Mises Stress")

    Dim outsetID As Long

    outsetID = oset.First

    'outsetID = ouSet.ID

    Dim sx As Double
    Dim sy As Double
    Dim txy As Double
    Dim d As Double
    Dim b As Double
    Dim p As Double
    Dim sig1 As Double
    Dim sig2 As Double
    Dim theta As Double
    Dim realvm As Double

    Dim Sigx As femap.Output
    Set Sigx = App.feOutput
    Dim Sigy As femap.Output
    Set Sigy =  App.feOutput
    Dim Tauxy As femap.Output
    Set Tauxy = App.feOutput
    Dim Delta As femap.Output
    Set Delta = App.feOutput
    Dim Beta As femap.Output
    Set Beta = App.feOutput
    Dim Phi As femap.Output
    Set Phi = App.feOutput
    Dim eqN As Double
    Dim eqD As Double
    Dim eqVonM As Double

    Dim RealVonMises As femap.Output
    Set RealVonMises = App.feOutput

    Dim totalElem As Long
    Dim elemCount As Long

    Dim feElem As femap.Elem
    Set feElem = App.feElem

    totalElem = App.Info_Count( FT_ELEM )

    Dim maxTop As Double
    Dim maxTopID As Long
    Dim maxBot As Double
    Dim maxBotID As Long

    Dim Pi As Double
    Pi = -Atn(-1)*4

    Dim DtoR As Double
    DtoR = (2*Pi)/360.0

    While outsetID

    maxTop = 0.0
    maxBot = 0.0


    Sigx.setID = outsetID
    rc = Sigx.Get(7020)
    Sigy.setID = outsetID
    rc = Sigy.Get(7021)
    Tauxy.setID = outsetID
    rc = Tauxy.Get(7023)
    Delta.setID = outsetID
    rc = Delta.Get(10007020)
    Beta.setID = outsetID
    rc = Beta.Get(10007021)
	Phi.setID = outsetID
	rc = Phi.Get(10007023)

    rc = RealVonMises.InitScalarAtElem( outsetID, 9000000, "Plate Top Max VonMises", FOT_STRESS, True )

    rc = App.feAppStatusShow( True, totalElem )

    elemCount = 0
    rc = feElem.Reset

    While feElem.Next
		elemCount = elemCount + 1
        rc = App.feAppStatusUpdate( elemCount )
		If feElem.type = FET_L_PLATE Then
			sx = Sigx.Value( feElem.ID )
    	    sy = Sigy.Value( feElem.ID )
	        txy= Tauxy.Value( feElem.ID )
        	b = Beta.Value( feElem.ID ) * DtoR
        	d = Delta.Value( feElem.ID ) * DtoR
        	p = Phi.Value( feElem.ID ) * DtoR

	        eqN = sx*sx*Sin(2*d) + sy*sy*Sin(2*b)+3*txy*txy*Sin(2*p)-(Abs(sx)*Abs(sy))*Sin(d+b)
    	    eqD = sx*sx*Cos(2*d)+sy*sy*Cos(2*b)+3*txy*txy*Cos(2*p)-Abs(sx)*Abs(sy)*Cos(d+b)

			If Abs( eqD ) > 0.0 Then

		        theta = Atn( -eqN / eqD )
    		    theta = theta / 2.0

	            eqVonM = Sqr( sx*sx*Cos(theta+d)*Cos(theta+d)+sy*sy*Cos(theta+b)*Cos(theta+b)+3*txy*txy*Cos(theta+p)*Cos(theta+p)-Abs(sx)*Abs(sy)*Cos(theta+d)*Cos(theta+b))

	            If eqVonM > maxTop Then
    	        	maxTop = eqVonM
        	        maxTopID = feElem.ID
            	End If

		        sx = sx * Cos( d-theta )
    		    sy = sy * Cos(b-theta )
        		txy = txy * Cos( p-theta )

		        sig1 = (( sx + sy ) / 2.0 ) + Sqr( (( sx - sy )/2.0)*((sx-sy)/2.0)+txy*txy)
    		    sig2 = (( sx + sy ) / 2.0 ) - Sqr( (( sx - sy )/2.0)*((sx-sy)/2.0)+txy*txy)

	    	    realvm = Sqr(((sig1-sig2)*(sig1-sig2)+sig2*sig2+sig1*sig1)/2.0)

		        RealVonMises.Value(feElem.ID) = eqVonM
		    End If
        End If
    Wend

    rc = App.feAppStatusShow( False, totalElem )

    rc = RealVonMises.Put( RealVonMises.ID )

    Sigx.setID = outsetID
    rc = Sigx.Get(7420)
    Sigy.setID = outsetID
    rc = Sigy.Get(7421)
    Tauxy.setID = outsetID
    rc = Tauxy.Get(7423)
    Delta.setID = outsetID
    rc = Delta.Get(10007420)
    Beta.setID = outsetID
    rc = Beta.Get(10007421)
	Phi.setID = outsetID
	rc = Phi.Get(10007423)

    rc = App.feAppStatusShow( True, totalElem )

    rc = RealVonMises.InitScalarAtElem( outsetID, 9000001, "Plate Bot Max VonMises", FOT_STRESS, True )

    rc = feElem.Reset
    elemCount = 0
    While feElem.Next
		elemCount = elemCount + 1
        rc = App.feAppStatusUpdate( elemCount )
        If feElem.type = FET_L_PLATE Then
			sx = Sigx.Value( feElem.ID )
    	    sy = Sigy.Value( feElem.ID )
	        txy= Tauxy.Value( feElem.ID )
        	b = Beta.Value( feElem.ID ) * DtoR
        	d = Delta.Value( feElem.ID ) * DtoR
        	p = Phi.Value( feElem.ID ) * DtoR

	        eqN = sx*sx*Sin(2*d) + sy*sy*Sin(2*b)+3*txy*txy*Sin(2*p)-(Abs(sx)*Abs(sy))*Sin(d+b)
    	    eqD = sx*sx*Cos(2*d)+sy*sy*Cos(2*b)+3*txy*txy*Cos(2*p)-Abs(sx)*Abs(sy)*Cos(d+b)

			If Abs( eqD ) > 0.0 Then

		        theta = Atn( -eqN / eqD )
    		    theta = theta / 2.0

	            eqVonM = Sqr( sx*sx*Cos(theta+d)*Cos(theta+d)+sy*sy*Cos(theta+b)*Cos(theta+b)+3*txy*txy*Cos(theta+p)*Cos(theta+p)-Abs(sx)*Abs(sy)*Cos(theta+d)*Cos(theta+b))

	            If eqVonM > maxBot Then
    	        	maxBot = eqVonM
        	        maxBotID = feElem.ID
            	End If


		        sx = sx * Cos( d-theta )
    		    sy = sy * Cos(b-theta )
        		txy = txy * Cos( p-theta )

		        sig1 = (( sx + sy ) / 2.0 ) + Sqr( (( sx - sy )/2.0)*((sx-sy)/2.0)+txy*txy)
    		    sig2 = (( sx + sy ) / 2.0 ) - Sqr( (( sx - sy )/2.0)*((sx-sy)/2.0)+txy*txy)

	    	    realvm = Sqr(((sig1-sig2)*(sig1-sig2)+sig2*sig2+sig1*sig1)/2.0)

		        RealVonMises.Value(feElem.ID) = eqVonM
		    End If
        End If
    Wend

    rc = RealVonMises.Put( RealVonMises.ID )

    Msg = "Maximum Equiv. Von Plate Top Von Mises Stress is " + Str$( maxTop ) + " at Elem " + Str$( maxTopID )
    rc = App.feAppMessage( FCM_NORMAL, Msg )
    Msg = "Maximum Equiv. Von Plate Bot Von Mises Stress is " + Str$( maxBot ) + " at Elem " + Str$( maxBotID )
    rc = App.feAppMessage( FCM_NORMAL, Msg )

    rc = App.feAppStatusShow( False, totalElem )

	outsetID = oset.Next

    'outsetID = ouSet.ID

    Wend
	
End Sub
