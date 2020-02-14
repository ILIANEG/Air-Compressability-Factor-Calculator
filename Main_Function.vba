Option Expicit 
Public Function CALCZ (P As Double, T As Double, PC As Double, TC As Double, W As Double) As Double
  Dim s_a As Double, s_b As Double, alph As Double, a As Double, b As DOuble
  s_a = (2955.119117 * (TC)^2)/(PC)
  s_b = (7.203658541 * (TC))/(PC)
  alph = (1 + (0.480+1.574*(W)-0.176*(W^2))*(1-((T)/(TC))^0.5))^2
  a = ((s_a)*(alph)*(P))/((6913.044464)*(T^2))
  b = ((s_b)*(P))/((83.14472)*(T))
  CALCZ = StartRidders()
End Function 
