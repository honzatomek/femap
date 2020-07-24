Rem File: Highlight_Tria3s.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elemSet As femap.Set
	Set elemSet = App.feSet

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	While feElem.Next
		If feElem.topology = FTO_TRIA3 Then
			elemSet.Add( feElem.ID)
		End If
	Wend

	If elemSet.count > 0 Then
		rc = App.feViewShow( FT_ELEM, elemSet.ID )

	End If

	
End Sub
