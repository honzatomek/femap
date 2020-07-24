Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim feElem As femap.Elem
    Set feElem = App.feElem

    Dim ndSet As femap.Set
    Set ndSet = App.feSet

    Dim elSet As femap.Set
    Set elSet = App.feSet

    DO_IT_AGAIN:

    rc = ndSet.Clear()
    rc = elSet.Clear()
    rc = App.feViewRegenerate( -1 )

    rc = feElem.SelectID( "Pick Seed Element" )
    If rc = -1 Then
		rc = ndSet.AddRule( feElem.ID, FGD_NODE_ONELEM )
		rc = elSet.AddSetRule( ndSet.ID, FGD_ELEM_BYNODE )
		rc = App.feViewShow( FT_ELEM, elSet.ID )
		GoTo DO_IT_AGAIN
    End If

End Sub
