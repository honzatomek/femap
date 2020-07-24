Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim el As femap.Elem
    Set el = App.feElem

    Dim eiD As Long
    Dim eiD2 As Long
    Dim eiD3 As Long
    Dim eiD4 As Long
    Dim eiD5 As Long
    Dim eiD6 As Long

    Dim sp1 As Long
    Dim sp2 As Long
    Dim sp3 As Long
    Dim sp4 As Long
    Dim sp5 As Long
    Dim sp6 As Long

	Dim st1 As Long
    Dim st2 As Long
    Dim st3 As Long
	Dim st4 As Long
    Dim st5 As Long
    Dim st6 As Long

    Dim end1 As Long
    Dim end2 As Long

    Dim p As femap.Prop
    Set p = App.feProp

	Dim eset As femap.Set
	Set eset = App.feSet

    eset.Select (FT_ELEM, True, "Select DOF Spring Elements")

    'rc = eset.AddRule( femap.FET_L_DOF_SPRING , FGD_ELEM_BYTYPE )

	Begin Dialog UserDialog 410,245,"Specify DOFs and Stiffness" ' %GRID:10,7,1,1
		CheckBox 40,21,90,14,"TX",.TX
		CheckBox 40,49,90,14,"TY",.TY
		CheckBox 40,77,90,14,"TZ",.TZ
		CheckBox 40,133,90,14,"RY",.RY
		CheckBox 40,161,90,14,"RZ",.RZ
		CheckBox 40,105,90,14,"RX",.RX
		OKButton 40,196,140,42
		CancelButton 220,196,150,42
		TextBox 240,14,140,21,.Val1
		TextBox 240,42,140,21,.Val2
		TextBox 240,70,140,21,.Val3
		TextBox 240,98,140,21,.Val4
		TextBox 240,126,140,21,.Val5
		TextBox 240,154,140,21,.Val6
		text 140,21,90,14,"Stiffness",.Stiff1
		text 140,49,90,14,"Stiffness",.Stiff2
		text 140,77,90,14,"Stiffness",.Stiff3
		text 140,105,90,14,"Stiffness",.Stiff4
		text 140,133,90,14,"Stiffness",.Stiff5
		text 140,161,90,14,"Stiffness",.Stiff6
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
	GoTo OK
	End If

    st1 = Val (dlg.Val1)
    st2 = Val (dlg.Val2)
    st3 = Val (dlg.Val3)
    st4 = Val (dlg.Val4)
    st5 = Val (dlg.Val5)
    st6 = Val (dlg.Val6)

    eiD = eset.First

    el.Get(eiD)

    If dlg.tx = 1 Then

    sp1 = el.propID
    p.type = FET_L_DOF_SPRING
   	p.flagI (2) = 1
    p.flagI (3) = 1
    p.pval (0) = st1
    p.Put (sp1)

    App.feAppMessage (FCM_WARNING, "Original DOF spring property, Property " + Str$(sp1) + " has been updated")

    End If

    If dlg.ty = 1 Then

    sp2 = p.NextEmptyID
    p.type = FET_L_DOF_SPRING
    p.flagI (2) = 2
    p.flagI (3) = 2
    p.pval (0) = st2
    p.Put (sp2)

    End If

    If dlg.tz = 1 Then

    sp3 = p.NextEmptyID
    p.type = FET_L_DOF_SPRING
    p.flagI (2) = 3
    p.flagI (3) = 3
    p.pval (0) = st3
    p.Put (sp3)

    End If

    If dlg.rx = 1 Then

    sp4 = p.NextEmptyID
    p.type = FET_L_DOF_SPRING
    p.flagI (2) = 4
    p.flagI (3) = 4
    p.pval (0) = st4
    p.Put (sp4)

    End If

    If dlg.ry = 1 Then

    sp5 = p.NextEmptyID
    p.type = FET_L_DOF_SPRING
    p.flagI (2) = 5
    p.flagI (3) = 5
    p.pval (0) = st5
    p.Put (sp5)

    End If

    If dlg.rz = 1 Then

    sp6 = p.NextEmptyID
    p.type = FET_L_DOF_SPRING
    p.flagI (2) = 6
    p.flagI (3) = 6
    p.pval (0) = st6
    p.Put (sp6)

    End If

    While eiD >0

    	el.Get (eiD)

		end1 = el.Node (0)
        end2 = el.Node (1)

        If dlg.tx = 1 Then

        el.type = FET_L_DOF_SPRING
        el.propID = sp1
        el.Put (eiD)

        End If

        If dlg.ty = 1 Then

        eiD2 = el.NextEmptyID
        el.type = FET_L_DOF_SPRING
        el.Node(0) = end1
        el.Node(1) = end2
        el.propID = sp2
        el.Put (eiD2)

        End If

        If dlg.tz = 1 Then

        eiD3 = el.NextEmptyID
        el.type = FET_L_DOF_SPRING
        el.Node(0) = end1
        el.Node(1) = end2
        el.propID = sp3
        el.Put (eiD3)

        End If

		If dlg.rx = 1 Then

        eiD4 = el.NextEmptyID
        el.type = FET_L_DOF_SPRING
        el.Node(0) = end1
        el.Node(1) = end2
        el.propID = sp4
        el.Put (eiD4)

        End If

        If dlg.ry = 1 Then

        eiD5 = el.NextEmptyID
        el.type = FET_L_DOF_SPRING
        el.Node(0) = end1
        el.Node(1) = end2
        el.propID = sp5
        el.Put (eiD5)

        End If

        If dlg.rz = 1 Then

        eiD6 = el.NextEmptyID
        el.type = FET_L_DOF_SPRING
        el.Node(0) = end1
        el.Node(1) = end2
        el.propID = sp6
        el.Put (eiD6)

        End If

		eiD = eset.Next
    Wend

    OK:

End Sub
