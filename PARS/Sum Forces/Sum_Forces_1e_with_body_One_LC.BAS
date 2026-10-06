' =============================================================================
' Description: Sums forces and moments for selected elements for one load case using legacy feCheckSumForces.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const DBUG As Boolean = False
Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_ELEM As Long = 8
Const C_FT_LOAD_DIR As Long = 12
Const C_FCM_NORMAL As Long = 0
Const C_FCM_WARNING As Long = 2
Const C_FCM_ERROR As Long = 3

Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long
    Dim expandGEOM As Boolean
    Dim doLIST As Boolean
    Dim useSETS As Boolean
    Dim nodeSET As Long
    Dim elemSET As Long
    Dim basePOINT(2) As Double
    Dim csysID As Long
    Dim summedFORCES(14) As Double

    er = 0

    Dim elset As Object
    Set elset = App.feSet()
    rc = elset.Select(C_FT_ELEM, True, "Select Elements for Force Summation:")
    If rc <> C_FE_OK Then GoTo Cleanup

    expandGEOM = True
    doLIST = True
    useSETS = True
    nodeSET = 0
    elemSET = elset.ID()
    basePOINT(0) = 0#
    basePOINT(1) = 0#
    basePOINT(2) = 0#
    csysID = 0

    rc = App.feCheckSumForces(expandGEOM, doLIST, useSETS, nodeSET, elemSET, basePOINT, csysID, summedFORCES)

Cleanup:
    If DBUG Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set elset = Nothing
    Set App = Nothing
End Sub
