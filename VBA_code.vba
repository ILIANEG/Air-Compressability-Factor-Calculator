Option Explicit
'Module variables announcment
Dim a As Double, b As Double, xl As Double, xu As Double

'Function that receives variables from user and defines variable "a" and "b", and returns Z'
Public Function CALCZ(P As Double, T As Double, PC As Double, TC As Double, W As Double) As Double
	Dim s_a As Double, s_b As Double, alph As Double
	s_a = (0.42747 * (TC) ^ 2) / (PC)
	s_b = (0.08664 * (TC)) / (PC)
	alph = (1 + (0.48 + 1.574 * (W) - 0.176 * (W ^ 2)) * (1 - ((T) / (TC)) ^ 0.5)) ^ 2
	a = ((s_a) * (alph) * (P)) / (T ^ 2)
	b = ((s_b) * (P)) / (T)
	'Calling solve() function which implements Ridder's algorythm'
	CALCZ = Solve()
End Function

'Function that calculates root using ridder's algorythm
Private Function Solve() As Double
	'Definition of required variables'
	Dim xm As Double, xr As Double, fxl As Double, fxm As Double, fxr As Double
	Dim fxu As Double, xrOld As Double, xrNew As Double, err As Double
	Dim firstCycle As Boolean
	firstCycle = True
	err = 100
	'Uses decrementing step to find upper intercept'
	Call findXl
	If xl <= 0 Then
		Solve = 0
	End If
	'While loop that runs Rider's algorythm until the seeking precision is reached
	Do Until err < 0.0001
		xm = (xl + xu) / 2
		'Calculates function of inputing x'
		fxl = calculateF(xl)
		fxm = calculateF(xm)
		fxu = calculateF(xu)
		xrNew = xm + (xm - xl) * ((fxl - fxu) / Abs(fxl - fxu) * fxm) / Sqr(fxm * fxm - fxl * fxu)
		If firstCycle = True Then
			xrOld = xrNew
			firstCycle = False
		Else
			err = Abs((xrNew - xrOld) / xrNew) * 100
			xrOld = xrNew
		End If
		'Next lines of the loop determines new xl and xu boundries for next round of ridders'
		If xm < xr Then
			If fxm * fxr < 0 Then
				xl = xm
				xu = xrNew
			ElseIf fxl * fxm < 0 Then
				xu = xm
			Else
				xl = xrNew
			End If
		Else
			If fxm * fxr < 0 Then
				xl = xrNew
				xu = xm
			ElseIf fxl * fxm < 0 Then
				xu = xrNew
			Else
				xl = xm
			End If
		End If
	Loop
	'Return Z'
	Solve = xrNew
End Function

'Evaluate function at point "x"'
Private Function calculateF(x As Double) As Double
	calculateF = x ^ 3 - x ^ 2 + x * (a - b - b ^ 2) - a * b
End Function

'* Find "biggest" root of the cubic function, rather then using incremental search we used decremental search'
'* Assumption of max value of Z = 0 was made'
Private Function findXl() As Double
	Dim step As Double, fxl As Double
	step = 0.0001
	xl = 2
	fxl = calculateF(xl)
	'Finds upper root by decrementing dtep'
	Do Until fxl * calculateF(xl) < 0
		xl = xl - step
	Loop
		xu = xl + step
End Function
