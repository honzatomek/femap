Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim am As femap.AnalysisMgr
    Set am = App.feAnalysisMgr

    Dim amID As Long

    'Choose the title

    am.title = "Static Analysis"

    'Choose the Solver

    am.Solver = FAM_NX_NASTRAN

    'Choose the analysis type

    am.AnalysisType = FAT_STATIC

    'Get the ID of the Active Constraint Set

    conset = App.Info_ActiveID (FT_BC_DIR)

    'If there is no Active Constraint Set, do not perform analysis

    If conset = 0 Then
    App.feAppMessageBox (0, "There is no active Constraint Set in this model.  No analysis will be performed.  Please activate a Constraint Set or create a new one and try again.")
    GoTo Fail
    End If

    'Set the Analysis Set Manager to use the Active Constraint Set.

    am.BCSet (0) = conset

    'Get the ID of the Active Load Set

    ldset = App.Info_ActiveID (FT_LOAD_DIR)

    'If there is no Active Load Set, do not perform analysis

	If ldset = 0 Then
    App.feAppMessageBox (0, "There is no active Load Set in this model. No analysis will be performed. Please activate a Load Set or create a new one and try again.")
    GoTo Fail
    End If

	'Set the Analysis Set Manager to use the Active Load Set

    am.BCSet (2) = ldset

    'Choose the output values to return.
    'Values in () correspond to values in the Output portion of section 5.3.1 of the API document
    '-1 after = signifies to request this output value for the entire model.
    '0 after = gets you nothing
    'values > 0 after = output requested for that group # only

    'Requests Nodal Applied Load - Default for Static Analysis in FEMAP
    am.Output(0) = -1
	'Requests Nodal Reaction Force - Default for Static Analysis in FEMAP
    am.Output(1) = -1
	'Requests Nodal Displacements - Default for Static Analysis in FEMAP
    am.Output(2) = -1
	'Requests Nodal Force Balance - Optional for Static Analysis...good for demos to show freebody
    am.Output(8) = -1
	'Requests Elemental Forces - Default for Static Analysis in FEMAP
    am.Output(15) = -1
	'Requests Elemental Stresses - Default for Static Analysis in FEMAP
    am.Output(16) = -1
    'Request 'Element Corner Results" -Default for Static Analysis in FEMAP
	am.CornerOutput = -1

    amID = am.NextEmptyID

    am.Put (amID)

    am.Analyze (amID)

    Fail:
	
End Sub
