Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    Dim useBodyLoad As Boolean
    Dim expandGEOM As Boolean
    Dim doLIST As Boolean
    Dim useSETS As Boolean
    Dim nodeSET As Long
    Dim elemSET As Long
    Dim loaddefSET As Long
    Dim basePOINT(2) As Double
    Dim csysID As Long
    Dim summedFORCES As Variant

    Dim elset As Object
    Set elset = App.feSet()

    rc = elset.Select(8, True, "Select Elements for Force Summation:")

    useBodyLoad = True
    expandGEOM = True
    doLIST = True
    useSETS = True
    nodeSET = 0
    elemSET = elset.ID()
    loaddefSET = 0
    basePOINT(0) = 0
    basePOINT(1) = 0
    basePOINT(2) = 0
    csysID = 0

    Dim ldsets As Object
    Set ldsets = App.feSet()
    rc = ldsets.AddAll(FT_LOAD_DIR)

    Dim LdSet As Object
    Set LdSet = App.feLoadSet()

    Dim activeLD As Long

    activeLD = App.Info_ActiveID(FT_LOAD_DIR)

    rc = ldsets.Reset()
    Do While ldsets.Next() <> 0
    	rc = LdSet.Get(ldsets.CurrentID())
		App.Info_ActiveID(FT_LOAD_DIR) = LdSet.ID()
		App.feAppMessage(FCM_ERROR, vbLf & "Check Sum Of Forces")
		App.feAppMessage(FCM_ERROR, "Load Set: [" & LdSet.ID() & "].." & LdSet.title())
		rc = App.feCheckSumForces2(useBodyLoad	, expandGEOM, doLIST, useSETS, nodeSET, elemSET, loaddefSET, basePOINT, csysID, summedFORCES)
    Loop

	App.Info_ActiveID(FT_LOAD_DIR) = activeLD

Cleanup:
	If DBUG Then Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set elset = Nothing
	Set LdSet = Nothing
	Set ldsets = Nothing
	Set App = Nothing
End Sub
