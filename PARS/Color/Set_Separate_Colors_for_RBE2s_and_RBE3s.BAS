' =============================================================================
' Description: Applies the PARS default colors: RBE2 color 72 and RBE3 color 122.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Default rigid-element colors used throughout PARS.
' RBE2: FEMAP palette 72 = RGB(128,255,128)
' RBE3: FEMAP palette 122 = RGB(128,255,255)
Const C_FE_OK As Long = -1
Const C_FET_RIGID As Long = 29
Const C_FGD_ELEM_BYTYPE As Long = 24
Const C_FCM_NORMAL As Long = 0
Const C_RBE2_COLOR As Long = 72
Const C_RBE3_COLOR As Long = 122

Sub Main
    Dim App As Object
    Dim rigidSet As Object
    Dim elObj As Object
    Dim rc As Long
    Dim eid As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Cleanup

    Set App = feGetObject()
    Set rigidSet = App.feSet()
    Set elObj = App.feElem()

    rc = rigidSet.AddRule(C_FET_RIGID, C_FGD_ELEM_BYTYPE)
    If rc = C_FE_OK Then
        rc = rigidSet.Reset()
        eid = rigidSet.Next()
        Do While eid <> 0
            rc = elObj.Get(eid)
            If rc = C_FE_OK Then
                If elObj.RigidInterpolate Then
                    elObj.color = C_RBE3_COLOR
                Else
                    elObj.color = C_RBE2_COLOR
                End If
                rc = elObj.Put(eid)
            End If
            eid = rigidSet.Next()
        Loop

        Call App.feAppMessage(C_FCM_NORMAL, "RBE2 color = 72; RBE3 color = 122.")
    End If

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    rc = App.feViewRegenerate(0)

    If errNumber <> 0 Then
        If Len(errDescription) > 100 Then errDescription = Left$(errDescription, 100)
        Call App.feAppMessage(3, "PARS Color error " & CStr(errNumber) & ": " & errDescription)
    End If

    Set elObj = Nothing
    Set rigidSet = Nothing
    Set App = Nothing
    On Error GoTo 0
End Sub
