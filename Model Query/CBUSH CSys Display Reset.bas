Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feGFXDelete (True, 0)

    App.feGFXReset

    App.feWindowRegenerate (0)
End Sub
