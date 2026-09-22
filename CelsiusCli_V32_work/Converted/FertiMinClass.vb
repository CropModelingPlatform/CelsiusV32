Option Strict Off
Option Explicit Off
Imports System
Public Class FertiMinClass
    Dim Nmin(731) As Double
    Dim Pmin(731) As Double
    Dim Kmin(731) As Double

    Dim rstDataFertiMin As ADODB.Recordset
    'VERSION 3 du 8-10 nov 2017'
    'choses à vérifier : commentaires insérés
    ' prévu pour bilans nutriments journaliers, pas opérationnel au 11/05/2026
    Public Sub LisFertiMinD(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass, GestionTechnique As GestionTechniqueClass)
    ' vérifier: manque les commentaires décrivant les variables
    Dim AnFertiMin As Integer
    Dim jourFertiMin As Integer
    Dim Nmineral As Double
    Dim Pmineral As Double
    Dim Kmineral As Double
    Dim Joursim As Integer
    Dim ErrFa As Boolean
    Dim Ndyear1 As Integer
    Dim msg As String
    On Error GoTo Err_LisFertiMinD

    If SimUnit.bSY_Bissextile Then Ndyear1 = 366 Else Ndyear1 = 365
    rstDataFertiMin = New ADODB.Recordset
    rstDataFertiMin.Open("SELECT * FROM FertiMin_List where IdTech_Com='" & SimUnit.sIdTec & "' Order by DateFertiMin", DataBase_Cnn, adOpenDynamic)
    Trouve = False
    '**** REDONDANT ! A verifier si critère = idsim ou idtec !!!
    While Not rstDataFertiMin.EOF And Not Trouve
    'attention idDclim si Escape, codeStat si mada
        If rstDataFertiMin("IdTech_Com") = SimUnit.sIdSim Then Trouve = True
        rstDataFertiMin.MoveNext
    End While
    If Not Trouve Then
        msg = "error in mineral fertilisation data table"
        GoTo ErrFA_LisFertiMinD
    End If
    'fin partie redondante a verifier
    rstDataFertiMin.MoveFirst

    While Not rstDataFertiMin.EOF

        AnFertiMin = rstDataFertiMin("YearFertiMin")
        jourFertiMin = rstDataFertiMin("JourYrFertiMin")
        Nmineral = rstDataFertiMin("Nmineral") 'vérifier unités utilisées
        Pmineral = rstDataFertiMin("Pmineral") 'vérifier unités utilisées
        Kmineral = rstDataFertiMin("Kmineral") 'vérifier unités utilisées

        If AnFertiMin = SimUnit.nStartYear Then
            Joursim = jourFertiMin - SimUnit.nStartDay + 1
        Else
            Joursim = jourFertiMin + Ndyear1 - SimUnit.nStartDay + 1
        End If
        If Joursim < 0 Then
            msg = "Attention, calendrier des apports minéraux incohérent avec calendrier de simulation"
            GoTo ErrFA_LisFertiMinD
        Else
            Nmin(Joursim) = Nmineral
            Pmin(Joursim) = Pmineral
            Kmin(Joursim) = Kmineral

        End If
        rstDataFertiMin.MoveNext
    End While
    rstDataFertiMin.Close
    rstDataFertiMin = Nothing



    Exit_LisFertiMinD:
        Exit Sub
    ErrFA_LisFertiMinD:
        ErrFa = True
    Console.Error.WriteLine(("ERREUR ! " & msg))
        Resume Exit_LisFertiMinD
    Err_LisFertiMinD:
    Console.Error.WriteLine(Err.Description)
        Resume Exit_LisFertiMinD
    End Sub
    Public Function dNmin(J As Integer) As Double
    Return Nmin(J)
    End Function
    Public Function dPmin(J As Integer) As Double
    Return Pmin(J)
    End Function
    Public Function dKmin(J As Integer) As Double
    Return Kmin(J)
    End Function
End Class
