Option Explicit

Public Function solve(A, B as Double) as Double
  Dim xl, xm, xu, fxl, fxm, fxu, xrOld, xrNew fxrNew, err as Double
  Dim counter as Boolean
  firstCycle = True 
  err = 100
  xl = 0
  xu = AB
  'while loop that runs Rider's algorythm until the seeking precision is reached
  Do Until err < 0.000001
    xm = (xl + xu) / 2
    fxl = Call calculateF(A, B, xl)
    fxm = Call calculateF(A, B, xm)
    fxu = Call calculateF(A, B, xu)
    xrNew =  xm + (xm - xl) * ((fxl - fxu)/Abs(fxl - fxu) * fxm)/Sqr(fxm * fxm - fxl * fxu)
    if firstCycle Then
      xrOld = xrNew
      firstCycle = False
    Else 
      err = Abs((xrNew - xrOld)/xrNew) * 100
      xrOld = xrNew
    End If
    if xm < xr then
      if fxm * fxr < 0 then
        xl = xm
        xu = xr
      ElseIf fxl * fxm < 0 then
        xu = xm
      Else
        xl = xr
      End If
    Else
      if fxm * fxr < 0 then
        xl = xr
        xu = xm
      ElseIf fxl * fxm < 0 then
        xu = xr
      Else
        xl = xm
      End If
    End If
  Loop
  solve = xr
End Function

'Function helps to abstract from our formula and avoid redundant copying of formula when calculating function of xl, xm, xu ad xr
Public Function calculateF(A, B, X as Double) as Double
  'Operator "^" is not used due to reported malfunctions on older versions of excel
  returnVal = X*X*X - X*X + X*(A - B - B*B) - A*B
End Function
