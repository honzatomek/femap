' =============================================================================
' Description: Test harness for the shared PARS math helper functions.
' Compatibility: Modern FEMAP BASIC only; depends on Math.bas.
' =============================================================================

Option Explicit On

'#Uses "Math.bas"

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Debug.Print MaxAbs(1#, -2#, 1#, 2#, 3#, 5#, 6#, 8#, -110#)

    Set App = Nothing
End Sub
