Attribute VB_Name = "Math"
Option Explicit On

Public Module Math
	Private er As Long

	'Ludolfs number pi
	Public Function pi() As Double
		'pi = 52163/16604
		pi = 3.14159265358979323846264338327950288419716939937510
	End Function

	'Eulers number
	Public Function e() As Double
		e = 2.71828182845904523536028747135266249775724709369995
	End Function

	'Logarithm of x with the base a
	Public Function LogA(ByVal x As Double, Optional ByVal a As Double = 10) As Double
		LogA = Log(x)/Log(a)
	End Function

	'Logarithm of x with the base 10
	Public Function Log10(ByVal x As Double) As Double
		Log10 = Log(x)/Log(10)
	End Function

	'Logarithm of x with the base e = 2,718
	Public Function LogN(ByVal x As Double) As Double
		LogN = Log(x)
	End Function

	'Tangent x
	Public Function tg(ByVal x As Double) As Double
		tg = Sin(x)/Cos(x)
	End Function

	'Cotangent x
	Public Function cotg(ByVal x As Double) As Double
		cotg = Cos(x)/Sin(x)
	End Function

	'Inverse sine x
	Public Function arcsin(ByVal x As Double) As Double
		arcsin = Atn(x/Sqr(-x*x+1))
	End Function

	'Inverse cosine x
	Public Function arccos(ByVal x As Double) As Double
		arccos = Atn(-x/Sqr(-x*x+1))+2*Atn(1)
	End Function

	'Inverse tangent x
	Public Function arctg(ByVal x As Double) As Double
		arctg = Atn(x)
	End Function

	'Inverse cotangent x
	Public Function arccotg(ByVal x As Double) As Double
		arccotg = 2 * Atn(1) - Atn(x)
	End Function

	'Angle between x axis and point (x,y) in 2D
	Public Function ArcTg2(ByVal x As Double, ByVal y As Double) As Double
	    Select Case x
	        Case Is > 0
	            ArcTg2 = Atn(y / x)
	        Case Is < 0
	            ArcTg2 = Atn(y / x) + pi() * Sgn(y)
	            If y = 0 Then ArcTg2 = ArcTg2 + pi()
	        Case Is = 0
	            ArcTg2 = pi()/2 * Sgn(y)
	    End Select
	End Function

	'Secant x
	Public Function sec(ByVal x As Double) As Double
		sec = 1/Cos(x)
	End Function

	'Cosecant x
	Public Function cosec(ByVal x As Double) As Double
		cosec = 1/Sin(x)
	End Function

	'Inverse secant x
	Public Function arcsec(ByVal x As Double) As Double
		arcsec = 2*Atn(1) - Atn(Sgn(x)/Sqr(x*x - 1))
	End Function

	'Inverse cosecant x
	Public Function arccosec(ByVal x As Double) As Double
		arccosec = Atn(Sgn(x)/Sqr(x*x-1))
	End Function

	'Hyperbolic sine x
	Public Function hsin(ByVal x As Double) As Double
		hsin = (Exp(x)-Exp(-x))/2
	End Function

	'Hyperbolic cosine x
	Public Function hcos(ByVal x As Double) As Double
		hcos = (Exp(x) + Exp(-x)) / 2
	End Function

	'Hyperbolic tangent x
	Public Function htg(ByVal x As Double) As Double
		htg = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
	End Function

	'Hyperbolic cotangent x
	Public Function hcotg(ByVal x As Double) As Double
		hcotg = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
	End Function

	'Hyperbolic secant x
	Public Function hsec(ByVal x As Double) As Double
		hsec = 2 / (Exp(x) + Exp(-x))
	End Function

	'Hyperbolic cosecant x
	Public Function hcosec(ByVal x As Double) As Double
		hcosec = 2 / (Exp(x) - Exp(-x))
	End Function

	'Inverse hyperbolix sine x
	Public Function harcsin(ByVal x As Double) As Double
		harcsin = Log(x + Sqr(x * x + 1))
	End Function

	'Inverse hyperbolix cosine x
	Public Function harccos(ByVal x As Double) As Double
		harccos = Log(x + Sqr(x * x - 1))
	End Function

	'Inverse hyperbolix tangent x
	Public Function harctg(ByVal x As Double) As Double
		harctg = Log((1 + x) / (1 - x)) / 2
	End Function

	'Inverse hyperbolix cotangent x
	Public Function harccotg(ByVal x As Double) As Double
		harccotg = Log((x+1)/(x-1))/2
	End Function

	'Inverse hyperbolix secant x
	Public Function harcsec(ByVal x As Double) As Double
		harcsec = Log((Sqr(-x * x + 1) + 1) / x)
	End Function

	'Inverse hyperbolix cosecant x
	Public Function harccosec(ByVal x As Double) As Double
		harccosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
	End Function

	'Maximum of values, max 12 variables
	Public Function Max(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
		Dim tmp() As Double
		Dim t As Double
		Dim temp As Variant
		Dim z As Long
		temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
		tmp = EvalArgs(temp)

		t = tmp(0)
		For z = 1 To UBound(tmp) Step 1
			If tmp(z) > t Then t = tmp(z)
		Next

		Max = t
	End Function

	'Minimum of values, max 12 variables
	Public Function Min(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
		Dim tmp() As Double
		Dim t As Double
		Dim temp As Variant
		Dim z As Long
		temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
		tmp = EvalArgs(temp)

		t = tmp(0)
		For z = 1 To UBound(tmp) Step 1
			If tmp(z) < t Then t = tmp(z)
		Next

		Min = t
	End Function

	'Maximum distance from 0 of a set of values, max 12 variables
	Public Function MaxAbs(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
		Dim tmp() As Double
		Dim t As Double
		Dim temp As Variant
		Dim z As Long
		temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
		tmp = EvalArgs(temp)

		t = tmp(0)
		For z = 1 To UBound(tmp) Step 1
			If Abs(tmp(z)) > Abs(t) Then t = tmp(z)
		Next

		MaxAbs = t
	End Function

	'Sum of values
	Public Function Sum(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
		Dim tmp() As Double
		Dim t As Double
		Dim temp As Variant
		Dim z As Long
		temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
		tmp = EvalArgs(temp)

		t = tmp(0)
		For z = 1 To UBound(tmp) Step 1
			 t = t + tmp(z)
		Next

		Sum = t
	End Function

	'Geometric average of values
	Public Function Avg(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
		Dim tmp() As Double
		Dim t As Double
		Dim temp As Variant
		Dim z As Long
		temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
		tmp = EvalArgs(temp)

		t = tmp(0)
		For z = 1 To UBound(tmp) Step 1
			 t = t + tmp(z)
		Next

		Avg = t / (UBound(tmp) + 1)
	End Function

	'Evaluation function for arguments - if they were passed or not
	Public Function EvalArgs(ByRef x() As Variant) As Variant
		Dim tmp() As Double
		Dim i As Long, j As Long, k As Long
		k = -1
		For i = 0 To UBound(x) Step 1
			If Not IsMissing(x(i)) Then
				If IsArray(x(i)) Then
					For j = 0 To UBound(x(i)) Step 1
						k = k + 1
						ReDim Preserve tmp(k)
						tmp(k) = CDbl(x(i)(j))
					Next
				Else
					k = k + 1
					ReDim Preserve tmp(k)
					tmp(k) = CDbl(x(i))
				End If
			End If
		Next
		EvalArgs = tmp
	End Function

End Module
