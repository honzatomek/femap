Rem File: importconnect.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	App.feRunCommand (1032,True)

	App.feFileProgramRun (False, True, False, "#SILENT(1){~1800}#DELAY(200)<A-S><OK><OK><Esc>")

End Sub
