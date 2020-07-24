Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim osetG As Object
    Set osetG = App.feOutputSet()

    Dim ovecG As Object
    Set ovecG = App.feOutput()

    Dim listcountG As Long
    Dim IDG As Variant
    Dim valueG As Variant

	Dim osets As Object
	Set osets = App.feSet()

    Dim oset As Object
    Set oset = App.feOutputSet()

    Dim ovec As Object
    Set ovec = App.feOutput()

    Dim listcount As Long
    Dim ID As Variant
    Dim value As Variant

    Dim g1 As Object
    Set g1 = App.feGroup()

    Dim g2() As Object
    'Set g2 = App.feGroup()

    Dim gset As Object
    Set gset = App.feSet()

'1
    If DBUG Then
    	rc = osetG.Get(11)
    	If rc <> -1 Then
    		er = 1
    		GoTo Cleanup
    	End If
    	Debug.Print osetG.title()
    Else
		rc = osets.SelectID(FT_OUT_CASE, "Select Dead Load Output Set:", i)
		If rc <> -1 Then
			er = 1
			GoTo Cleanup
		End If
		rc = osetG.Get(i)
		If rc <> -1 Then
			er = 1
			GoTo Cleanup
		End If
		rc = osets.Clear()
	End If

'2
	rc = osets.SelectMultiID(FT_OUT_CASE, 1, "Select Output Sets other than Dead Load:")
	If rc <> -1 Then
		er = 2
		GoTo Cleanup
	End If
	ReDim g2(osets.Count())

'3
	rc = ovecG.GetFromSet(osetG.ID, 3022)
	If rc <> -1 Then
		er = 3
		GoTo Cleanup
	End If
	rc = ovecG.GetOutputList(listcountG, IDG, valueG)
	g1.title = "Temp for Output processing"
	For i = 0 To listcountG - 1 Step 1
		If valueG(i) < 0 Then
			rc = g1.Add(FT_ELEM, IDG(i))
		End If
	Next
	rc = g1.Put(g1.NextEmptyID())
	rc = gset.Add(g1.ID)

'4
	rc = osets.Reset()
	i = -1
	Do While osets.Next() <> 0
		i = i + 1
		Set g2(i) = App.feGroup()
		g2(i).title = "Temp for Output processing"
		rc = oset.Get(osets.CurrentID())
		If DBUG Then Debug.Print oset.title()
		rc = ovec.GetFromSet(oset.ID, 3022)
		If rc <> -1 Then
			er = 4
			GoTo Cleanup
		End If
		rc = ovec.GetOutputList(listcount, ID, value)
		value = AddVector(valueG, value)
		For j = 0 To listcountG - 1 Step 1
			If value(j) < 0 Then
				rc = g2(i).Add(FT_ELEM, IDG(j))
			End If
		Next
		rc = g2(i).Put(g2(i).NextEmptyID())
		rc = gset.Add(g2(i).ID)
	Loop

'5
	rc = App.feGroupBoolean2(gset.ID, 0, FGB_INALL, 0, "ELEMENTS IN COMPRESSION")
	If rc <> -1 Then
		er = 5
		GoTo Cleanup
	End If

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
		For i = UBound(g2) To 0 Step -1
			Set g2(i) = Nothing
		Next
		If Not DBUG Then Call App.feDelete(FT_GROUP, gset.ID)
		Set g1 = Nothing
		Set gset = Nothing
		Set ovec = Nothing
		Set ovecG = Nothing
		Set oset = Nothing
		Set osetG = Nothing
		Set osets = Nothing
		Set App = Nothing
End Sub

Private Function AddVector(ByRef A() As Double, ByRef B() As Double) As Variant
	Dim C() As Double
	Dim i As Long
	If (LBound(A) = LBound(B)) And (UBound(A) = UBound(B)) Then
		ReDim C(LBound(A) To UBound(A))
		For i = LBound(A) To UBound(A) Step 1
			C(i) = A(i) + B(i)
		Next
	Else
		AddVector = 0
		Exit Function
	End If
	AddVector = C
End Function

Private Function SubtractVector(ByRef A() As Double, ByRef B() As Double) As Variant
	Dim C() As Double
	Dim i As Long
	If (LBound(A) = LBound(B)) And (UBound(A) = UBound(B)) Then
		ReDim C(LBound(A) To UBound(A))
		For i = LBound(A) To UBound(A) Step 1
			C(i) = A(i) - B(i)
		Next
	Else
		SubtractVector = 0
		Exit Function
	End If
	SubtractVector = C
End Function
