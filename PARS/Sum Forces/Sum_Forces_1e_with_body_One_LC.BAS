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

    rc = App.feCheckSumForces2(useBodyLoad	, expandGEOM, doLIST, useSETS, nodeSET, elemSET, loaddefSET, basePOINT, csysID, summedFORCES)

Cleanup:
	If DBUG Then Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set elset = Nothing
	Set App = Nothing
End Sub
