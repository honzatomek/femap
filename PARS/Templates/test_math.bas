Option Explicit On

'#Uses "Math.bas"

Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Debug.Print Math.MaxAbs(1, -2, Array(1,2,3), 5, 6,8, -110)
    
End Sub
