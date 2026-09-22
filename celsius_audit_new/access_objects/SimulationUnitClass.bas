Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
' Classe originale FA
Dim SY_Bissextile As Boolean
Dim IdSim As String
Dim IdTec As String
Dim IdWeather As String
Dim idSoil As String
Dim idWeedCom As String
Dim idbiotic_Alea As String
Dim idIni As String
Dim idGenParam As String
Dim idCodModel As Integer
Dim StartYear As Integer
Dim StartDay As Integer
Dim EndYear As Integer
Dim EndDay As Integer
Dim EndDOY As Integer 'Jour de fin de simulation comptés à partir du 1er janvier de l'année de début
Dim NbJourSimul As Integer
Dim CodCC As String
'
'VERSION 3 du 8-10 nov 2017
'Modif FA du 30/08/18 pour scenarios CC: lecture nouveau champ: codCC
'modif FA du 18/12/21 recursivité avec codesuite lu dans simunitlist
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026

Sub ReadSimUnitParameters(rstSimulation As ADODB.Recordset)

IdSim = rstSimulation!IdSim
IdTec = rstSimulation!IdTech_Com
IdWeather = rstSimulation!IdWeather
idSoil = rstSimulation!idSoil
idIni = rstSimulation!idIni
idGenParam = rstSimulation!idGenParam
idCodModel = rstSimulation!idCodModel
StartYear = rstSimulation!StartYear
StartDay = rstSimulation!StartDay
EndYear = rstSimulation!EndYear
EndDay = rstSimulation!EndDay
SY_Bissextile = ((StartYear Mod 4) = 0)
CodCC = rstSimulation!CodCC
Codesuite = rstSimulation!Codesuite
'calcul de EndDoY prévu pour un cas général où on simule n années sans réinitialiser l'état
'du systeme (cas pas opérationnel dans le reste du projet, seuls cas valables: EndYear-Start Year=0 ou 1
' et EndDOY < 366+365 soit  731)
EndDOY = EndDay + ((EndYear - StartYear) Mod 2) * (365 - CInt(SY_Bissextile))
NbJourSimul = EndDOY - StartDay + 1
End Sub
Property Get sIdSim() As String
sIdSim = IdSim
End Property
Property Get sIdTec() As String
sIdTec = IdTec
End Property
Property Get sIdWeather() As String
sIdWeather = IdWeather
End Property
Property Get sIdSoil() As String
sIdSoil = idSoil
End Property

Property Get sidGenParam() As String
sidGenParam = idGenParam
End Property
Property Get nStartYear() As Integer
nStartYear = StartYear
End Property
Property Get nStartDay() As Integer
nStartDay = StartDay
End Property
Property Get nEndYear() As Integer
nEndYear = EndYear
End Property
Property Get nEndDay() As Integer
nEndDay = EndDay
End Property
Property Get bSY_Bissextile() As Boolean
bSY_Bissextile = SY_Bissextile
End Property
Property Get nEndDOY() As Integer
nEndDOY = EndDOY
End Property
Property Get nidCodModel() As Integer
nidCodModel = idCodModel
End Property
Property Get nNbJourSimul() As Integer
nNbJourSimul = NbJourSimul
End Property
Property Get sidIni() As String
sidIni = idIni
End Property
Property Get sCodcc() As String
sCodcc = CodCC
End Property
Property Let nNbJourSimul(NJsimul As Integer)
NbJourSimul = NJsimul
End Property