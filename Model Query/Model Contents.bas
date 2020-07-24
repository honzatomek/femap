Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim emp As Long
	Dim ent As Variant
	App.feAppModelContents( True, emp, ent )
End Sub
