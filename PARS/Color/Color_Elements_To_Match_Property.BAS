' =============================================================================
' Description: Sets each element color to the color of its referenced property.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1

Sub Main
    Dim App As Object
    Dim elObj As Object
    Dim prop As Object
    Dim rc As Long
    Dim lastProp As Long

    Set App = feGetObject()
    Set elObj = App.feElem()
    Set prop = App.feProp()

    lastProp = -1
    rc = elObj.First()
    Do While rc = C_FE_OK
        If elObj.propID > 0 Then
            If lastProp <> elObj.propID Then
                lastProp = elObj.propID
                rc = prop.Get(lastProp)
            End If
            If rc = C_FE_OK Then
                elObj.color = prop.color
                rc = elObj.Put(elObj.ID)
            End If
        End If
        rc = elObj.Next()
    Loop

    rc = App.feViewRegenerate(0)
    Set prop = Nothing
    Set elObj = Nothing
    Set App = Nothing
End Sub
