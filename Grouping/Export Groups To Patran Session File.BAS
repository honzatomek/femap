'#Language "WWB-COM"

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim fName As String

    Dim gr As femap.Group
    Set gr = App.feGroup

    Dim grSet As femap.Set
    Set grSet = App.feSet

    Dim ndelSet As femap.Set
    Set ndelSet = App.feSet
	Dim sType As String
	Dim i As Long
	Dim grTitle As String
	Dim sEOL As String
	Dim iStart As Long
	Dim iThis As Long
	Dim iStop As Long
	Dim nDigits As Long
	Dim sID As String

    If grSet.SelectMultiID( FT_GROUP, 1, "Select Groups to Write to Patran Session File" ) = FE_OK Then
		If App.feFileGetName( "Select File to Write", "Patran Session Files (*.SES)", "*.SES", False, fName ) = FE_OK Then
			Open fName For Output As #1

			Print #1, "$# Session file"
			Print #1, "$# FEMAP Groups to Patran Session File Converter"
			Print #1, "$# Original FEMAP Model File " + App.ModelName
			Print #1, "sys_poll_option( 2 )

			nDigits = Len(LTrim$(Str$(grSet.Last())))

			While gr.NextInSet( grSet.ID ) = FE_OK

				sID =Trim$(Str$(gr.ID))
				While Len(sID) < nDigits
					sID = "0" + sID
				Wend
				grTitle = sID+"_"+gr.title
				Print #1, "ga_group_create( """+grTitle+""" )"

				For i = 1 To 2
					ndelSet.Clear()
					If  i = 1 Then
						sType = "Node"
						If gr.List( FGR_NODE ) IsNot Nothing Then
							ndelSet.AddSet( gr.List( FGR_NODE ).ID  )
						End If
					Else
						sType = "Element"
						If gr.List( FGR_ELEM ) IsNot Nothing Then
							ndelSet.AddSet( gr.List( FGR_ELEM ).ID  )
						End If
					End If

					If  ndelSet.IsEmpty = FE_FAIL Then

							Print #1, "ga_group_entity_add( """+grTitle+""", @"
							Print #1, """ "+sType+" "" // @"

							ndelSet.Reset()
							Do
								iStart = ndelSet.Next()
								iThis   = iStart

								Do
									iStop = ndelSet.NextAfter( iThis )
									If  iStop = FE_FAIL Then
										GoTo WRITE_IT
									ElseIf iStop <> iThis+1 Then
WRITE_IT:
									    If iStop = FE_FAIL Then
									    	sEOL = """ )"
									    Else
									    	sEOL = """ // @"
									    End If
										If  iThis = iStart Then
											Print #1, """ "+Trim$(Str$(iStart))+sEOL
										Else
											Print #1,""" "+Trim$(Str$(iStart))+":"+Trim$(Str$(iThis))+sEOL
										End If
										iStart = iStop
										iThis = iStart
									Else
										iThis = iStop
									End If
								Loop Until iStart = FE_FAIL
							Loop Until iStart = FE_FAIL
					End If
				Next

			Wend
			Print #1, "sys_poll_option( 0 )"
			Print #1, "$# Session file"
			Close #1
		End If
    End If

    App.feAppMessage( FCM_WARNING, "Session File "+fName+" Written." )
End Sub
