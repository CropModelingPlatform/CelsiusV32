Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'VX: Variable explicative, VE: variable d'état, VS: variable simulée,PX: paramètre
Dim CO2c As Integer 'Vx Concentration en CO2 de l'atmosphere (ppm)
Dim ddat As Date
Dim Tmoy(0 To 731) As Double 'VX
Dim Tmax(0 To 731) As Double 'VX
Dim Tmin(0 To 731) As Double 'VX
Dim Rg(0 To 731) As Double 'VX
Dim Etp(0 To 731) As Double 'VX
Dim Plu(0 To 731) As Double 'VX
Dim DOY(0 To 731) As Integer 'jour de l'année
Dim DAP(0 To 731) As Integer 'jour après semis / plantation (day after planting)
Dim CurrentYear(0 To 731) As Integer
Dim Altitude As Integer 'VX
Dim Latitude As Double 'VX
Dim DAYL(0 To 731) As Double 'Vx longueur astronomique du jour (h)
Dim rstDataClim As ADODB.Recordset
Dim Ndyear1 As Integer
Dim ConcNplu As Double 'VX concentration en N des précipitations UNITE a compléter
Dim TMS As Double 'VX temperature moyenne de la saison de culture, utile pour k2 minéralisation
Dim NprecipTot As Double 'VX total de l'azote apporté par les précipitations (cumul sur la période de simulation)en kg/ha/mm
'VERSION 3 du 8-10 nov 2017'
' introduction teneur en N des pluies
' modifs FA du 24/11/2021 calcul de TMS
' modifs FA du 6/1/22 lecture tableau de CO2 en fonction des années
' modifs FA du 11/08/22 calcul des apports annuels de'azote par précipitations à partir de concentration en N des pluies


Sub LisClimD(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)
Dim Complete As Boolean
Dim LackClim As Boolean
Dim Trouve As Boolean
Dim ErrFa As Boolean
Dim n As Integer

Dim NJsimul As Integer
Dim msg As String
Dim yearCO2 As Integer 'année du tableau des concentrations en CO2
Dim CO2year As Double 'CO2 de l'année

On Error GoTo Err_LisClimD
Set rstDataClim = New ADODB.Recordset
'lecture des données non temporelles de la station dans la table ListPannexes
rstDataClim.Open "ListPAnnexes", DataBase_Cnn, adOpenDynamic
rstDataClim.MoveFirst
Trouve = False
While Not rstDataClim.EOF And Not Trouve

If rstDataClim!idDclim = SimUnit.sIdWeather Then Trouve = True

    rstDataClim.MoveNext
Wend
If Not Trouve Then
    msg = "error in weather station list table"
    GoTo ErrFA_LisClimD
End If
rstDataClim.MovePrevious

Latitude = rstDataClim!latitudeDD
Altitude = rstDataClim!Altitude
CO2c = rstDataClim!CO2c
ConcNplu = rstDataClim!ConcNplu

rstDataClim.Close
' à faire éventuellement
' vérifier la cohérence des années de simulation avec les années présentes dans la table

' lecture de [CO2] annuelles dans une table si codcc est à "0"
If SimUnit.sCodcc = "0" Then
  Set rstDataClim = New ADODB.Recordset
  rstDataClim.Open "CO2Yearly", DataBase_Cnn, adOpenDynamic
  Trouve = False
  rstDataClim.MoveFirst
  While Not rstDataClim.EOF And Not Trouve
  yearCO2 = rstDataClim!yearCO2
  If yearCO2 = SimUnit.nStartYear Then
      Trouve = True
      CO2year = rstDataClim!CO2
  End If
  rstDataClim.MoveNext
  Wend
  If Not Trouve Then
      msg = "error in [CO2] table (CO2Yearly): data for simulation year not found, simulation run with average [CO2] read in ListPAnnexes"
      MsgBox (msg)
  Else
   CO2c = CO2year
  End If
  rstDataClim.Close
End If


Set rstDataClim = New ADODB.Recordset

'lecture des données journalieres de la station dans la table Dweather
' important de faire le tri croissant sur date dans la lecture de la table (order by) !
rstDataClim.Open "SELECT * FROM Dweather where idDclim='" & SimUnit.sIdWeather & "'AND annee>=" & SimUnit.nStartYear & " Order by annee, Jda", DataBase_Cnn, adOpenDynamic

Complete = False
LackClim = False
Trouve = False

While Not rstDataClim.EOF And Not Trouve

    If rstDataClim!idDclim = SimUnit.sIdWeather And rstDataClim!annee = SimUnit.nStartYear Then Trouve = True
    rstDataClim.MoveNext
Wend
If Not Trouve Then
    msg = "error in climatic data table"
    GoTo ErrFA_LisClimD
End If


rstDataClim.MoveFirst
n = 1
' n est dans le calendrier de simulation (n augmente à partir du jour de début de simulation
While (Not Complete) And (Not rstDataClim.EOF) And (Not LackClim)

CurrentYear(n) = rstDataClim!annee
DOY(n) = rstDataClim!Jda


If (DOY(n) >= SimUnit.nStartDay Or CurrentYear(n) = SimUnit.nStartYear + 1) Then

'introduction ici (avant transformation de DOY ds le calendrier de 0 à 731) du calcul de la durée du jour
    DAYL(n) = LenDay(Latitude, DOY(n))

    If SimUnit.bSY_Bissextile Then Ndyear1 = 366 Else Ndyear1 = 365
    If DOY(n) <= DOY(n - 1) Then DOY(n) = DOY(n) + Ndyear1

    Tmin(n) = rstDataClim!Tmin
    Tmax(n) = rstDataClim!Tmax
    Tmoy(n) = Tmoy_TnTx(Tmin(n), Tmax(n))
    Rg(n) = rstDataClim!Rg
    Etp(n) = rstDataClim!Etp
    Plu(n) = rstDataClim!Plu
    NprecipTot = NprecipTot + Plu(n) * ConcNplu
    LackClim = IsNull(Tmin(n) + Tmax(n))
    
    If DOY(n) = SimUnit.nEndDOY Then Complete = True
    n = n + 1
End If
rstDataClim.MoveNext

Wend
'fermeture table et libération mémoire de l'objet
rstDataClim.Close
Set rstDataClim = Nothing


NJsimul = n - 1
If LackClim Or Not Complete Then
    msg = SimUnit.sIdSim
    If LackClim Then msg = msg & " Caution ! Missing climatic data. Simulation performed until first missing data"
    If Not Complete Then msg = msg & "Fin de simulation apres fin des données climatiques !!! Réduisez la période de simulation"
    MsgBox (msg)
    NJsimul = NJsimul - 1
    SimUnit.nNbJourSimul = NJsimul
End If

Exit_LisClimD:
    Exit Sub
ErrFA_LisClimD:
    ErrFa = True
    MsgBox ("ERREUR ! " & msg)
    Resume Exit_LisClimD
Err_LisClimD:
    MsgBox Err.Description
    Resume Exit_LisClimD
End Sub
Sub update_DAP(idebut As Integer, iplt As Integer)
'calcul des jours après semis / plantation et stockage dans tableau
' n est dans le calendrier de simulation (n=1 pour la date de début de simulation)
Dim n As Integer
Dim nbjouran As Integer
For n = 0 To 731
    DAP(n) = idebut - iplt + n - 1

Next n
End Sub
Sub CorrigTAlt(AltiCult As Integer, NJsimul As Integer)
Dim i As Integer
For i = 1 To NJsimul

 Tmin(i) = Tmin(i) - (AltiCult - Altitude) * 0.6 / 100
 Tmax(i) = Tmax(i) - (AltiCult - Altitude) * 0.6 / 100
 Tmoy(i) = Tmoy_TnTx(Tmin(i), Tmax(i))
Next i
End Sub
Sub CorrigTSerres(Joursim As Integer)
' modèle empirique FA/ Jenny Montagne

Const SERA As Double = 0.24 'constantes empiriques Jenny Montagne
Const SERB As Double = 0.095

Tmin(Joursim) = Tmax(Joursim) - SERB * (Tmax(Joursim) + Tmin(Joursim)) + (SERA * (1 - SERB) * Rg(Joursim))
Tmax(Joursim) = Tmax(Joursim) + SERA * Rg(Joursim)
Tmoy(Joursim) = Tmoy_TnTx(Tmin(Joursim), Tmax(Joursim))
 
End Sub
Sub TempMoySaison(DurCycMax As Integer, NJsimul As Integer)
Dim n As Integer, i As Integer
TMS = 0
For i = 1 To NJsimul
If DAP(i) > -30 And DAP(i) < DurCycMax Then
    TMS = TMS + Tmoy(i)
    n = n + 1
End If
Next i
TMS = TMS / n
End Sub
Function LenDay(lat As Double, J As Integer) As Double
' calcul de la durée astronomique du jour
'attention, VBasic utilise le passage d'arguments Byref par défaut, et la variable DataClim.Latitude est modifiée par cette fonction
' si on ne passe pas par la variable intérmédiaire locale "latrad". On aurait pê pu spécifier aussi "Byval" et laisser lat=pi*lat/180, à vérifier
Dim Declin As Double, SolarAngle As Double, latrad As Double
Const pi = 3.14159265
latrad = pi * lat / 180
Declin = 0.409 * Sin((2 * pi * J / 365) - 1.39)
SolarAngle = Tan(latrad) * Tan(Declin)

SolarAngle = -Atn(-SolarAngle / Sqr(-SolarAngle * SolarAngle + 1)) + 2 * Atn(1)
' note: Atn(1)= pi/4
LenDay = 24 * SolarAngle / pi
End Function
Function Tmoy_TnTx(Tmin As Double, Tmax As Double) As Double
Tmoy_TnTx = (Tmin + Tmax) / 2
End Function
Property Get dTmin(J As Integer) As Double
dTmin = Tmin(J)
End Property
Property Get dTmax(J As Integer) As Double
dTmax = Tmax(J)
End Property
Property Get dTmoy(J As Integer) As Double
dTmoy = Tmoy(J)
End Property
Property Get dRg(J As Integer) As Double
dRg = Rg(J)
End Property
Property Get dEtp(J As Integer) As Double
dEtp = Etp(J)
End Property
Property Get dPlu(J As Integer) As Double
dPlu = Plu(J)
End Property
Property Get nDAP(J As Integer) As Integer
nDAP = DAP(J)
End Property
Property Get nDOY(J As Integer) As Integer
nDOY = DOY(J)
End Property
Property Get nCurrentYear(J As Integer) As Integer
nCurrentYear = CurrentYear(J)
End Property
Property Get dDAYL(J As Integer) As Double
dDAYL = DAYL(J)
End Property
Property Get nCO2c() As Integer
nCO2c = CO2c
End Property
Property Get dConcNplu() As Double
dConcNplu = ConcNplu
End Property
Property Get dTMS() As Double
dTMS = TMS
End Property
Property Get dNprecipTot() As Double
dNprecipTot = NprecipTot
End Property