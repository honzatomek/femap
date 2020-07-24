'Written bij Alexander Naatje, Femto Engineering
'http://www.femto.nl

Sub Main
Dim App As femap.model
Set App = feFemap()

Dim rc As Variant
Dim i As Integer
Dim viewID As Long
Dim vi As Object
Set vi = App.feView

rc = App.feAppGetActiveView(viewID)

rc = vi.Get(viewID)

For  i=0 To 2
   If  (vi.rotation(i) < -135) Then
     vi.rotation(i) = -180
   Else
     If  (vi.rotation(i) < -45) Then
       vi.rotation(i) = -90
     Else
       If  (vi.rotation(i) < 45) Then
         vi.rotation(i) = 0
       Else
         If  (vi.rotation(i) < 135) Then
           vi.rotation(i) = 90
         Else
             vi.rotation(i) = 180
           End If
       End If
     End If
  End If
Next i

rc = vi.Put(viewID)

rc = vi.Regenerate()

End Sub
