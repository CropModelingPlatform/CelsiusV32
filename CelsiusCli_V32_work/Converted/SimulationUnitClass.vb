Option Strict Off
Option Explicit Off
Imports System
Public Class SimulationUnitClass
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
    Public Sub ReadSimUnitParameters(rstSimulation As ADODB.Recordset)

    IdSim = rstSimulation("IdSim")
    IdTec = rstSimulation("IdTech_Com")
    IdWeather = rstSimulation("IdWeather")
    idSoil = rstSimulation("idSoil")
    idIni = rstSimulation("idIni")
    idGenParam = rstSimulation("idGenParam")
    idCodModel = rstSimulation("idCodModel")
    StartYear = rstSimulation("StartYear")
    StartDay = rstSimulation("StartDay")
    EndYear = rstSimulation("EndYear")
    EndDay = rstSimulation("EndDay")
    SY_Bissextile = ((StartYear Mod 4) = 0)
    CodCC = rstSimulation("CodCC")
    Codesuite = rstSimulation("Codesuite")
    'calcul de EndDoY prévu pour un cas général où on simule n années sans réinitialiser l'état
    'du systeme (cas pas opérationnel dans le reste du projet, seuls cas valables: EndYear-Start Year=0 ou 1
    ' et EndDOY < 366+365 soit  731)
    EndDOY = EndDay + ((EndYear - StartYear) Mod 2) * (365 - CInt(SY_Bissextile))
    NbJourSimul = EndDOY - StartDay + 1
    End Sub
    Public Function sIdSim() As String
    Return IdSim
    End Function
    Public Function sIdTec() As String
    Return IdTec
    End Function
    Public Function sIdWeather() As String
    Return IdWeather
    End Function
    Public Function sIdSoil() As String
    Return idSoil
    End Function

    Public Function sidGenParam() As String
    Return idGenParam
    End Function
    Public Function nStartYear() As Integer
    Return StartYear
    End Function
    Public Function nStartDay() As Integer
    Return StartDay
    End Function
    Public Function nEndYear() As Integer
    Return EndYear
    End Function
    Public Function nEndDay() As Integer
    Return EndDay
    End Function
    Public Function bSY_Bissextile() As Boolean
    Return SY_Bissextile
    End Function
    Public Function nEndDOY() As Integer
    Return EndDOY
    End Function
    Public Function nidCodModel() As Integer
    Return idCodModel
    End Function
    Public Property nNbJourSimul As Integer
        Get
            Return NbJourSimul
        End Get
        Set(value As Integer)
            NbJourSimul = value
        End Set
    End Property

    Public Function sidIni() As String
    Return idIni
    End Function
    Public Function sCodcc() As String
    Return CodCC
    End Function
End Class
