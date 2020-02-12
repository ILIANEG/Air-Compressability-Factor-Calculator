Option Explicit

Public Function solve(A As Double, B As Double) As Double
  Dim xl As Double, xm As Double, xu As Double, xr As Double, fxl As Double, fxm As Double, fxr As Double, fxu As Double, xrOld As Double, xrNew As Double, fxrNew As Double, err As Double
  Dim firstCycle As Boolean
  firstCycle = True
  err = 100
  xl = 0
  xu = 2
  'while loop that runs Rider's algorythm until the seeking precision is reached
  Do Until err < 0.0000000001
    xm = (xl + xu) / 2
    fxl = calculateF(A, B, xl)
    fxm = calculateF(A, B, xm)
    fxu = calculateF(A, B, xu)
    xrNew = xm + (xm - xl) * ((fxl - fxu) / Abs(fxl - fxu) * fxm) / Sqr(fxm * fxm - fxl * fxu)
    If firstCycle = True Then
      xrOld = xrNew
      firstCycle = False
    Else
      err = Abs((xrNew - xrOld) / xrNew) * 100
      xrOld = xrNew
    End If
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
  solve = xrNew
End Function

'Function helps to abstract from our formula and avoid redundant copying of formula when calculating function of xl, xm, xu ad xr
Private Function calculateF(A As Double, B As Double, X As Double) As Double
  'Operator "^" is not used due to reported malfunctions on older versions of excel
  calculateF = X * X * X - X * X + X * (A - B - B * B) - A * B
End Function

Private Function findXu(A, B)
  Dim xu As Double, Dim fxu As Double
  xu = 0
  Do Until fxu > 0
    xu = xu + 0.5
    fxu = calculateF(A, B, fxu)
  Loop
  findXu = xu
