Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim tRadius As Double
    tRadius = 0.0

    Dim feSurf As femap.Surface
    Set feSurf = App.feSurface

    Dim surfSET As femap.Set
    Set surfSET = App.feSet

    Dim feGroup As femap.Group
    Set feGroup = App.feGroup

    Dim testRadius As Double
    Dim vBase As Variant
    Dim vAxis As Variant

    rc = App.feGetReal( "Enter Target Radius Value", 0.0, 100.0, tRadius )

    If tRadius > 0 Then
    	While feSurf.Next = -1
			If feSurf.IsCylinder() = -1 Then
				rc = feSurf.Cylindrical( testRadius, vBase, vAxis )
				If (Abs( tRadius - testRadius ) / tRadius) < 0.01 Then
					rc = surfSET.Add( feSurf.ID )
				End If
			End If
    	Wend
    	If surfSET.Count > 0 Then
    		feGroup.title = "Surfaces with Radius = " + Str$( tRadius )
    		rc = feGroup.SetAdd( FT_SURFACE, surfSET.ID )
    		rc = feGroup.Put( feGroup.NextEmptyID )
    		rc = App.feViewRegenerate( 0 )
	   	End If
    End If
End Sub
