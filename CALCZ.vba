Option Explicit
'Module variables announcment
Dim a As Double, b As Double, xl As Double

Public Function CALCZ(P As Double, T As Double, PC As Double, TC As Double, W As Double) As Double
  Dim s_a As Double, s_b As Double, alph As Double
  s_a = (2955.119117 * (TC) ^ 2) / (PC)
  s_b = (7.203658541 * (TC)) / (PC)
  alph = (1 + (0.48 + 1.574 * (W) - 0.176 * (W ^ 2)) * (1 - ((T) / (TC)) ^ 0.5)) ^ 2
  a = ((s_a) * (alph) * (P)) / ((6913.044464) * (T ^ 2))
  b = ((s_b) * (P)) / ((83.14472) * (T))
  CALCZ = Solve()
End Function

'Function that calculates root using ridder's algorythm
Private Function Solve() As Double
  Dim xm As Double, xu As Double, xr As Double, fxl As Double, fxm As Double, fxr As Double, fxu As Double, xrOld As Double, xrNew As Double, fxrNew As Double, err As Double
  Dim firstCycle As Boolean
  firstCycle = True
  err = 100
  xu = findXu()
  xl = 
  'while loop that runs Rider's algorythm until the seeking precision is reached
  Do Until err < 1E-06
    xm = (xl + xu) / 2
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
  Solve = xrNew
End Function

'Function helps to abstract from our formula and avoid redundant copying of formula when calculating function of xl, xm, xu ad xr
Private Function calculateF(x As Double) As Double
  'Operator "^" is not used due to reported malfunctions on older versions of excel
  calculateF = x * x * x - x * x + x * (a - b - b * b) - a * b
End Function

Private Function findXu() As Double
  Dim xu As Double, fxu As Double, step As Double
  xu = 0
  fxu = calculateF(xu)
  fxl = calculateF(0)
  step = 1
  Do Until fxu > 0 and fxl * fxu < 0
    xu = xu + step
    fxu = calculateF(xu)
    If fxl * fxu > 0 Then
      xu = xu - step
      step = step / 2
  Loop
    findXu = xu
    xl = xu - step
End Function
