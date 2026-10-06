' =============================================================================
' Description: Creates/manages standard orthogonal and isometric FEMAP views using legacy view APIs.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit
'Creates Views in +XY, -XY, +XZ, -XZ, +YZ, -YZ and axonometry

Const DBUG As Boolean = False
Const C_FT_VIEW As Long = 22
Const C_FE_FAIL As Long = 0

Private rotX(6) As Double
Private rotY(6) As Double
Private rotZ(6) As Double
Private viewNames(6) As String
Private curview As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long

    On Error GoTo ExitHere

    Call InitViews()
    curview = App.Info_ActiveID(C_FT_VIEW)

    Begin Dialog UserDialog 340,147,"Views Manager",.viewmanfunc ' %GRID:10,7,1,1
        GroupBox 10,7,210,133,"Select View:",.GroupBox1
        PushButton 20,28,90,21,"+XY",.PushButton1
        PushButton 120,28,90,21,"-XY",.PushButton2
        PushButton 20,56,90,21,"+XZ",.PushButton3
        PushButton 120,56,90,21,"-XZ",.PushButton4
        PushButton 20,84,90,21,"+YZ",.PushButton5
        PushButton 120,84,90,21,"-YZ",.PushButton6
        PushButton 20,112,190,21,"Axonometry",.PushButton7
        CancelButton 230,112,100,21
        PushButton 230,28,100,21,"Create Views",.create
        PushButton 230,56,100,21,"Reset Views",.rset
    End Dialog
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then rc = C_FE_FAIL

ExitHere:
    On Error Resume Next
    Set App = Nothing
End Sub

Private Sub InitViews()
    rotX(0) = 0#:    rotY(0) = 0#: rotZ(0) = 0#
    rotX(1) = 180#:  rotY(1) = 0#: rotZ(1) = 0#
    rotX(2) = -90#:  rotY(2) = 0#: rotZ(2) = 0#
    rotX(3) = -90#:  rotY(3) = 0#: rotZ(3) = 180#
    rotX(4) = -90#:  rotY(4) = 0#: rotZ(4) = -90#
    rotX(5) = -90#:  rotY(5) = 0#: rotZ(5) = 90#
    rotX(6) = -75#:  rotY(6) = 0#: rotZ(6) = -135#

    viewNames(0) = "+XY"
    viewNames(1) = "-XY"
    viewNames(2) = "+XZ"
    viewNames(3) = "-XZ"
    viewNames(4) = "+YZ"
    viewNames(5) = "-YZ"
    viewNames(6) = "axonometry"
End Sub

Private Function viewmanfunc(DlgItem, Action, SuppValue) As Boolean
    Select Case Action
    Case 1
    Case 2
        viewmanfunc = True
        Select Case DlgItem
        Case "Cancel"
            viewmanfunc = False
        Case "create"
            Call CreateViews()
            viewmanfunc = False
        Case "rset"
            Call ResetViews()
            viewmanfunc = False
        Case "PushButton1"
            Call RotateCurrent(rotX(0), rotY(0), rotZ(0)): viewmanfunc = False
        Case "PushButton2"
            Call RotateCurrent(rotX(1), rotY(1), rotZ(1)): viewmanfunc = False
        Case "PushButton3"
            Call RotateCurrent(rotX(2), rotY(2), rotZ(2)): viewmanfunc = False
        Case "PushButton4"
            Call RotateCurrent(rotX(3), rotY(3), rotZ(3)): viewmanfunc = False
        Case "PushButton5"
            Call RotateCurrent(rotX(4), rotY(4), rotZ(4)): viewmanfunc = False
        Case "PushButton6"
            Call RotateCurrent(rotX(5), rotY(5), rotZ(5)): viewmanfunc = False
        Case "PushButton7"
            Call RotateCurrent(rotX(6), rotY(6), rotZ(6)): viewmanfunc = False
        End Select
    Case 3
    Case 4
    Case 5
    Case 6
    End Select
End Function

Private Sub RotateCurrent(ByVal rx As Double, ByVal ry As Double, ByVal rz As Double)
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long
    Dim v As Object
    Set v = App.feView()

    On Error GoTo ExitHere

    rc = v.Get(curview)
    If rc = C_FE_FAIL Then rc = v.Get(1)
    If rc = C_FE_FAIL Then GoTo ExitHere

    rc = v.AlignCSys(0)
    rc = v.Rotate(0, 0, 0, 1, 0, 0, rx)
    rc = v.Rotate(0, 0, 0, 0, 1, 0, ry)
    rc = v.Rotate(0, 0, 0, 0, 0, 1, rz)
    rc = v.AutoscaleAll(True)
    rc = v.Put(v.ID)
    rc = v.Redraw()

ExitHere:
    On Error Resume Next
    Set v = Nothing
    Set App = Nothing
End Sub

Private Sub CreateViews()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long
    Dim id As Long
    Dim exists As Boolean
    Dim v As Object
    Dim viewSet As Object
    Set v = App.feView()
    Set viewSet = App.feSet()

    On Error GoTo ExitHere
    Call InitViews()

    For i = 0 To 6
        exists = False
        rc = viewSet.Clear()
        rc = viewSet.AddAll(C_FT_VIEW)
        id = viewSet.First()
        Do While id <> C_FE_FAIL
            rc = v.Get(id)
            If rc <> C_FE_FAIL Then
                If v.title = viewNames(i) Then
                    exists = True
                    Exit Do
                End If
            End If
            id = viewSet.Next()
        Loop

        If Not exists Then
            rc = v.Get(1)
            If rc = C_FE_FAIL Then GoTo ExitHere
            v.title = viewNames(i)
            rc = v.AlignCSys(0)
            rc = v.Rotate(0, 0, 0, 1, 0, 0, rotX(i))
            rc = v.Rotate(0, 0, 0, 0, 1, 0, rotY(i))
            rc = v.Rotate(0, 0, 0, 0, 0, 1, rotZ(i))
            rc = v.AutoscaleAll(True)
            rc = v.Put(v.NextEmptyID())
        End If
    Next
    Call App.feViewRedraw(0)

ExitHere:
    On Error Resume Next
    Set viewSet = Nothing
    Set v = Nothing
    Set App = Nothing
End Sub

Private Sub ResetViews()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long
    Dim id As Long
    Dim v As Object
    Dim viewSet As Object
    Set v = App.feView()
    Set viewSet = App.feSet()

    On Error GoTo ExitHere
    Call InitViews()

    rc = viewSet.AddAll(C_FT_VIEW)
    For i = 0 To 6
        id = viewSet.First()
        Do While id <> C_FE_FAIL
            rc = v.Get(id)
            If rc <> C_FE_FAIL Then
                If v.title = viewNames(i) Then
                    rc = v.AlignCSys(0)
                    rc = v.Rotate(0, 0, 0, 1, 0, 0, rotX(i))
                    rc = v.Rotate(0, 0, 0, 0, 1, 0, rotY(i))
                    rc = v.Rotate(0, 0, 0, 0, 0, 1, rotZ(i))
                    rc = v.AutoscaleAll(True)
                    rc = v.Put(v.ID)
                    Exit Do
                End If
            End If
            id = viewSet.Next()
        Loop
    Next
    Call App.feViewRedraw(0)

ExitHere:
    On Error Resume Next
    Set viewSet = Nothing
    Set v = Nothing
    Set App = Nothing
End Sub
