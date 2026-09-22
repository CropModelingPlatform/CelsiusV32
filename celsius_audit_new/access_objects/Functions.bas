Option Compare Database
' modifié 2/08/13 (latrad au lieu de lat ligne latrad = pi * lat / 180 ?? - commentaire entre () ajouté le 24/04/14)
'modif FA le 24/04/2014 fonction lenday lat --> latrad ligne SolarAngle = Tan(latrad) * Tan(Declin)

Function Julian(dat1) As Variant
If IsNull(dat1) Then
Julian = Null
Else
Julian = 1 + DateDiff("d", "01/01/" & CStr(Year(dat1)), dat1)
End If
End Function
Function LenDay(lat As Double, J As Integer) As Double
Dim Declin As Double, SolarAngle As Double
Dim latrad As Double

Const pi As Double = 3.14159265
latrad = pi * lat / 180
Declin = 0.409 * Sin((2 * pi * J / 365) - 1.39)
SolarAngle = Tan(latrad) * Tan(Declin)
SolarAngle = -Atn(-SolarAngle / Sqr(-SolarAngle * SolarAngle + 1)) + 2 * Atn(1)
' note: Atn(1)= pi/4
LenDay = 24 * SolarAngle / pi
End Function
Function Ra(lat, J) As Double
'returns Solar radiation at the top of the atmosphere as a function of day of the year and latitude
' ra in Mj/m2/day

Dim Dr As Double, Declin As Double, SolarAngle As Double
Dim latrad As Double
Const pi = 3.14159265

latrad = pi * lat / 180
Dr = 1 + 0.033 * Cos(2 * pi * J / 365)
Declin = 0.409 * Sin((2 * pi * J / 365) - 1.39)
SolarAngle = Tan(latrad) * Tan(Declin)
If SolarAngle ^ 2 >= 1 Then SolarAngle = 0.99999
SolarAngle = -Atn(-SolarAngle / Sqr(-SolarAngle * SolarAngle + 1)) + 2 * Atn(1)
Ra = (SolarAngle * Sin(latrad) * Sin(Declin)) + (Cos(latrad) * Cos(Declin) * Sin(SolarAngle))
Ra = 24 * 60 * 0.082 * Dr * Ra / pi
End Function
Function ET0pm(lat, Alt, J, Tn, Tx, Tm, Un, Ux, Vm, Rg) As Variant
'Calculates ET0 according to FAO Penman-Monteith Equation (Bull.FAO#56)
'altitude Alt in m,latitude in decimal degrees
'J in number of the day in the year
'Tn, Tx and Tm respectively minimal,maximal and mean daily temperature in degrees C,
'Un, Ux, respectively minimal and maximal relative humidity of air in %
'Vm,average wind speed per day  in m/s,
'Rg global radiation in MJ/m2/day
' All variables assumed measured at 2m above soil

Dim adv As Double, Rad As Double, gamma As Double, E0SatTn As Double, E0SatTx As Double, SlopeSat As Double, Ea As Double
Dim VPD As Double, Rso As Double, Rns As Double, Rnl As Double

Const lambda = 2.45, sigma = 0.000000004903
'jeu test (exemple du bull FAO56 - Allen et al., pp72-73: resultat=3.9 (ici 3.88) et Ra=41.09 _fonction ci-dessus)
'lat = 50.8
'Alt = 100
'Tn = 12.3
'Tx = 21.5
'Tm = 16.9
'Ux = 84
'Un = 63
'Vm = 2.078
'Rg = 22.07
'J = 187
'fin jeu test

'terme advectif
If IsNull(lat + Alt + J + Tn + Tx + Tm + Un + Ux + Vm + Rg) Then
    ET0pm = Null
Else
  Un = Un / 100
  Ux = Ux / 100
  gamma = 101.3 * ((293 - 0.0065 * (Alt)) / 293) ^ 5.26
  gamma = 0.000665 * gamma
  E0SatTn = 0.6108 * Exp(17.27 * Tn / (Tn + 237.3))
  E0SatTx = 0.6108 * Exp(17.27 * Tx / (Tx + 237.3))
  SlopeSat = 4098 * (0.6108 * Exp(17.27 * Tm / (Tm + 237.3))) / ((Tm + 237.3) ^ 2)
  Ea = 0.5 * ((E0SatTn * Ux) + (E0SatTx * Un))
  VPD = ((E0SatTn + E0SatTx) / 2) - Ea
  adv = gamma * 900 * Vm * VPD / (Tm + 273)
'terme radiatif
      'shortwave
  Rso = (0.00002 * Alt + 0.75) * Ra(lat, J)
  Rns = (1 - 0.23) * Rg
       'longwave
  Rnl = Rg / Rso
  If Rnl > 1 Then Rnl = 1
  Rnl = (Rnl * 1.35 - 0.35) * (-0.14 * Sqr(Ea) + 0.34)
  Tn = Tn + 273.16
  Tx = Tx + 273.16
  Rnl = sigma * (Tn ^ 4 + Tx ^ 4) * Rnl / 2
        'radiation balance assuming soil heat flux is 0 at a day time step
  Rad = 0.408 * SlopeSat * (Rns - Rnl)
'        ajout des deux termes
  ET0pm = (Rad + adv) / (SlopeSat + gamma * (0.34 * Vm + 1))
End If
End Function
Function ET0pen48(lat, Alt, J, Tn, Tx, Tm, Un, Ux, Vm, Rg) As Variant
'Calculates ET0 according to original Penman Equation (1948)
'''' attention: not thouroughly checked
Dim adv As Double, Rad As Double, gamma As Double, E0SatTn As Double, E0SatTx As Double, SlopeSat As Double
Dim Ea As Double, VPD As Double, Rso As Double, Rns As Double, Rnl As Double
  Un = Un / 100
  Ux = Ux / 100
  gamma = 101.3 * ((293 - 0.0065 * (Alt)) / 293) ^ 5.26
  gamma = 0.000665 * gamma
  E0SatTn = 0.6108 * Exp(17.27 * Tn / (Tn + 237.3))
  E0SatTx = 0.6108 * Exp(17.27 * Tx / (Tx + 237.3))
  SlopeSat = 4098 * (0.6108 * Exp(17.27 * Tm / (Tm + 237.3))) / ((Tm + 237.3) ^ 2)
  Ea = 0.5 * ((E0SatTn * Ux) + (E0SatTx * Un))
  VPD = ((E0SatTn + E0SatTx) / 2) - Ea
  
  adv = gamma * 900 * Vm * VPD / (Tm + 273)
'terme radiatif
      'shortwave
  Rso = (0.00002 * Alt + 0.75) * Ra(lat, J)
  Rns = (1 - 0.23) * Rg
       'longwave
  Rnl = Rg / Rso
  If Rnl > 1 Then Rnl = 1
  Rnl = (Rnl * 1.35 - 0.35) * (-0.14 * Sqr(Ea) + 0.34)
 Tn = Tn + 273.16
 Tx = Tx + 273.16
 Rnl = sigma * (Tn ^ 4 + Tx ^ 4) * Rnl / 2
        'radiation balance assuming soil heat flux is 0 at a day time step
  Rad = 0.408 * SlopeSat * (Rns - Rnl)


ET0pen48 = Rad + adv
End Function

Function ET0pm_Tdew(lat, Alt, J, Tn, Tx, Tm, Tdewn, Tdewx, Vm, Rg) As Variant
'Calculates ET0 according to FAO Penman-Monteith Equation (Bull.FAO#56)
'altitude Alt in m, latitude in decimal degrees
'J in number of the day in the year
'Tn, Tx and Tm respectively minimal,maximal and mean daily temperature in degrees C,
'Tdewn, Tdewx, dewpoint temperature(in °C) at respectively minimal and maximal temperature
'Vm,average wind speed per day in m/s,
'Rg global radiation in MJ/m2/day
' All variables assumed measured at 2m above soil

Dim adv As Double, Rad As Double, gamma As Double, E0SatTn As Double, E0SatTx As Double, SlopeSat As Double, Ea As Double
Dim VPD As Double, Rso As Double, Rns As Double, Rnl As Double

Const lambda = 2.45, sigma = 0.000000004903


'terme advectif
If IsNull(lat + Alt + J + Tn + Tx + Tm + Tdewn + Tdewx + Vm + Rg) Then
    ET0pm_Tdew = Null
Else
  gamma = 101.3 * ((293 - 0.0065 * (Alt)) / 293) ^ 5.26
  gamma = 0.000665 * gamma
  E0SatTn = 0.6108 * Exp(17.27 * Tn / (Tn + 237.3))
  E0SatTx = 0.6108 * Exp(17.27 * Tx / (Tx + 237.3))
  SlopeSat = 4098 * (0.6108 * Exp(17.27 * Tm / (Tm + 237.3))) / ((Tm + 237.3) ^ 2)
  'Ea = 0.5 * ((E0SatTn * Ux) + (E0SatTx * Un))
 ' remplacé par ligne suivante puisque Tdewpoint (equ 12 Bull FAO 56, p36)
  Ea = 0.5 * 0.6108 * (Exp(17.27 * Tdewx / (Tdewx + 237.3)) + Exp(17.27 * Tdewn / (Tdewn + 237.3)))
  VPD = ((E0SatTn + E0SatTx) / 2) - Ea
  adv = gamma * 900 * Vm * VPD / (Tm + 273)
'terme radiatif
      'shortwave
  Rso = (0.00002 * Alt + 0.75) * Ra(lat, J)
  Rns = (1 - 0.23) * Rg
       'longwave
  Rnl = Rg / Rso
  If Rnl > 1 Then Rnl = 1
  Rnl = (Rnl * 1.35 - 0.35) * (-0.14 * Sqr(Ea) + 0.34)
  Tn = Tn + 273.16
  Tx = Tx + 273.16
  Rnl = sigma * (Tn ^ 4 + Tx ^ 4) * Rnl / 2
        'radiation balance assuming soil heat flux is 0 at a day time step
  Rad = 0.408 * SlopeSat * (Rns - Rnl)
'        ajout des deux termes
  ET0pm_Tdew = (Rad + adv) / (SlopeSat + gamma * (0.34 * Vm + 1))
End If
End Function