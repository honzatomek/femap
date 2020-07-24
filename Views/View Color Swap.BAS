Rem File: ViewColorSwap.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim viewID As Long
	Dim feView As femap.View
	Set feView = App.feView

	rc = App.feAppGetActiveView( viewID )

	rc = feView.Get(viewID)

	If feView.WindowShadeBack Then
		'background is shaded, standard FEMAP, must flip things
		feView.WindowBackColor = FCL_WHITE
		'turns shaded background off
		feView.WindowShadeBack = False
		'Label Parameter color
		feView.color(FVI_LABEL) = FCL_BLACK
		'View Legend - View Color
		feView.color(FVI_VIEW_LEGEND) = FCL_BLACK
		'Post Titles - View Color
		feView.color(FVI_POST_TITLES) = FCL_BLACK
		'Contour Legend - Label Color = 1..Use View Color
		feView.ColorMode (FVI_CONTOUR_LEGEND) = 1
		'Contour Legend - View Color
		feView.color(FVI_CONTOUR_LEGEND) = FCL_BLACK
		'XY Axes Style - Axes Color
		feView.color(FVI_XY_AXES_STYLE) = FCL_BLACK
		'XY Titles - View Color
        feView.color(FVI_XY_TITLES) = FCL_BLACK
       	'XY X Range/Grid - Grid Color
        feView.color(FVI_XY_XAXIS) = FCL_BLACK
       	'XY Y Range/Grid - Grid Color
        feView.color(FVI_XY_YAXIS) = FCL_BLACK

	Else
		feView.WindowBackColor = FCL_BLACK
		feView.WindowShadeBack = True
		feView.color(FVI_LABEL) = FCL_WHITE
		feView.color(FVI_VIEW_LEGEND) = FCL_WHITE
		feView.color(FVI_POST_TITLES) = FCL_WHITE
		'Contour Legend - Label Color = 0..Contour Colors
		feView.ColorMode (FVI_CONTOUR_LEGEND) = 0
		feView.color(FVI_XY_AXES_STYLE) = FCL_WHITE
        feView.color(FVI_XY_TITLES) = FCL_WHITE
        feView.color(FVI_XY_XAXIS) = FCL_BLUE
        feView.color(FVI_XY_YAXIS) = FCL_BLUE

	End If

	rc = feView.Put( viewID )
	rc = App.feViewRegenerate( 0 )

End Sub
