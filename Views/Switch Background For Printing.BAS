Rem File: Switch Background for Printing.bas

Dim App As femap.model

Dim feView As Object
Dim savedView As Object
Dim savedViewTitle As String
Dim viewID As Long
Dim actViewID As Long

Sub LoadOView( act, save )
  rc = savedView.Put(act)
  rc = savedView.Visible(True)
  rc = App.feAppSetActiveView(act)
  rc = feView.Delete(save)
  rc = savedView.Regenerate()
End Sub

Sub Main
  Set App = feFemap()
  Set feView = App.feView()
  Set savedView = App.feView()

  Set stdBackColor2 = 80
  Set stdBackColor = 0
  Set stdFGColor = 124
  Set prtBackColor = 124
  Set prtBackColor2 = 0
  Set prtFGColor = 0
  tempViewTitle="femapAPISavedPrintView"

  'Create view objects
  Dim num As Long
  Dim rc As Long
  Dim index As Long
  Dim e As Long

  'Get the Active View ID
  rc = App.feAppGetActiveView(viewID)

  If rc = FE_OK Then
	rc = feView.Get(viewID)
	If rc = FE_OK Then
	  actViewID = viewID
'=============== revert view ==================
	  'If the view color is already white, switch back to saved
	  If feView.WindowBackColor = prtBackColor Then
		savedView.Reset()
		e = FE_OK
		Do While e <> FE_FAIL
		  e = savedView.Next()
		  If e <> FE_FAIL Then
			If(StrComp(savedView.title, tempViewTitle) = 0) Then
			  savedView.title = feView.title
			  savedViewID = savedView.ID
			  LoadOView( actViewID, savedViewID)
			  Exit All
			End If
		  End If
		Loop
		' if orig view not found switch back to femap standard
		'================= reset view =================
		feView.WindowBackColor2 =stdBackColor2
		feView.WindowBackColor = stdBackColor
		feView.WindowShadeBack = True
		feView.WindowShadeBackOpt = 0
		feView.color(FVI_LABEL) = stdFGColor
		feView.color(FVI_VIEW_LEGEND) = stdFGColor
		feView.color(FVI_POST_TITLES) = stdFGColor
		feView.ColorMode(FVI_CONTOUR_LEGEND) = 0
		feView.color(FVI_CONTOUR_LEGEND) = stdFGColor
		feView.color(FVI_SHRINK_ELEM) = stdElemColor
		feView.color(FVI_NODE) = stdFGColor
		feView.color(FVI_ELEMENT) = stdFGColor
		feView.Draw(FVI_ORIGIN) = 0
		rc = feView.Visible(True)
		rc = feView.Put(viewID)
		rc = feView.Regenerate()

'======== switch view to white bg And contrasting fg ========
	  Else         '( white BG , black FG )
		Dim origTitle As String
		' save orig view as temp view
		rc = savedView.Get(viewID)
		If (rc = FE_OK ) Then
		  origTitle = savedView.title
		  num = savedView.NextEmptyID()
		  savedView.title = tempViewTitle
		  savedView.Put(num)
		End If
		' get original back into feview
		feView.Get(viewID)
		' reset title to original
		feView.title = origTitle
		' change colors
		feView.WindowBackColor2 = prtBackColor
		feView.WindowBackColor = prtBackColor
		feView.WindowShadeBack = False
		feView.ColorMode(FVI_LABEL) = 1
		feView.color(FVI_LABEL) =prtFGColor
		feView.color(FVI_VIEW_LEGEND) =prtFGColor
		feView.color(FVI_POST_TITLES) = prtFGColor
		'feView.ColorMode(FVI_NODE) = 1
		' feView.color(FVI_NODE) = prtFGColor
		If( feView.Draw(FVI_FILL_ELEM) ) Then
		  feView.ColorMode(FVI_FILLED_EDGES) = 3
		  feView.color(FVI_FILLED_EDGES ) = prtFGColor
		Else
		  feView.ColorMode(FVI_ELEMENT) = 1
		  feView.color(FVI_ELEMENT) = prtFGColor
		End If
		feView.ColorMode(FVI_SHRINK_ELEM) = 1
		feView.color(FVI_SHRINK_ELEM) = prtFGColor
		feView.ColorMode (FVI_CONTOUR_LEGEND) = 1
		feView.color(FVI_CONTOUR_LEGEND) = prtFGColor
		feView.color(FVI_XY_AXES_STYLE) = 0
		feView.color(FVI_XY_TITLES) = prtFGColor
		feView.color(FVI_XY_XAXIS) = prtFGColor
		feView.color(FVI_XY_YAXIS) = prtFGColor
		feView.Draw(FVI_ORIGIN) = False
		feView.Draw(FVI_VIEW_AXES) = True

		rc = feView.Put(viewID)
		rc= feView.Visible(True)
		rc = feView.Regenerate()
	  End If
	End If
  End If
End Sub

