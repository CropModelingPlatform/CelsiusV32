Option Strict Off
Option Explicit Off
Imports System
Public Class FertiOrgaClass
    Dim Norg(731) As Double
    Dim Porg(731) As Double
    Dim Korg(731) As Double
    Dim Qorga(731) As Double

    Dim Norganique As Double
    Dim Porganique As Double
    Dim Korganique As Double
    Dim QorgaTot As Double
    Dim TypeMorga As String
    Dim CsurNRes As Double

    ' attention fonctionnel seulement si un seul apport par simunit, d'un seul type et quantité
    Dim rstDataFertiOrg As ADODB.Recordset
    'VERSION 3 du 8-10 nov 2017'
    ' choses à vérifier voir commentaires insérés
    ' prévu pour bilans nutriments journaliers, pas opérationnel au 11/05/2026
    Public Sub LisFertiOrgD(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass, GestionTechnique As GestionTechniqueClass)
    'commentaires décrivant les variables à insérer
    Dim AnFertiMin As Integer
    Dim jourFertiOrg As Integer


    Dim Joursim As Integer
    Dim ErrFa As Boolean
    Dim Ndyear1 As Integer
    Dim msg As String
    On Error GoTo Err_LisFertiOrgD

    If SimUnit.bSY_Bissextile Then Ndyear1 = 366 Else Ndyear1 = 365
    rstDataFertiOrg = New ADODB.Recordset
    rstDataFertiOrg.Open("SELECT * FROM FertiOrg_List where IdTech_Com='" & SimUnit.sIdTec & "' Order by DateFertiOrg", DataBase_Cnn, adOpenDynamic)
    ' *** à vérifier : redondant avec ouverture requête, voir si critère = idsim ou idtec
    Trouve = False

    While Not rstDataFertiOrg.EOF And Not Trouve
    'attention idDclim si Escape, codeStat si mada
        If rstDataFertiOrg("IdTech_Com") = SimUnit.sIdSim Then Trouve = True
        rstDataFertiOrg.MoveNext
    End While
    If Not Trouve Then
        msg = "error in organic fertilisation data table"
        GoTo ErrFA_LisFertiOrgD
    End If
    ' fin bloc redondant avec ouverture requête, voir si critère = idsim ou idtec
    rstDataFertiOrg.MoveFirst

    While Not rstDataFertiOrg.EOF

        AnFertiOrg = rstDataFertiOrg("YearFertiOrg")
        jourFertiOrg = rstDataFertiOrg("JourYrFertiOrg")
        Norganique = rstDataFertiOrg("Norganique") 'vérifier unités
        Porganique = rstDataFertiOrg("Porganique") 'vérifier unités
        Korganique = rstDataFertiOrg("Korganique") 'vérifier unités
        QorgaTot = rstDataFertiOrg("QorgaTot")
        TypeMorga = rstDataFertiOrg("TypeMorga")
        CsurNRes = rstDataFertiOrg("CsurNRes")

        If AnFertiOrg = SimUnit.nStartYear Then
            Joursim = jourFertiOrg - SimUnit.nStartDay + 1
        Else
            Joursim = jourFertiOrg + Ndyear1 - SimUnit.nStartDay + 1
        End If
        If Joursim < 0 Then
            msg = "Attention, calendrier des apports organiques incohérent avec calendrier de simulation"
            GoTo ErrFA_LisFertiOrgD
        Else
            Norg(Joursim) = Norganique
            Porg(Joursim) = Porganique
            Korg(Joursim) = Korganique
            Qorga(Joursim) = QorgaTot
        End If
        rstDataFertiOrg.MoveNext
    End While
    rstDataFertiOrg.Close



    Trouve = False
    rstDataFertiOrg = Nothing


    Exit_LisFertiOrgD:
        Exit Sub
    ErrFA_LisFertiOrgD:
        ErrFa = True
    Console.Error.WriteLine(("ERREUR ! " & msg))
        Resume Exit_LisFertiOrgD
    Err_LisFertiOrgD:
    Console.Error.WriteLine(Err.Description)
        Resume Exit_LisFertiOrgD
    End Sub
    Public Function dNorg(J As Integer) As Double
    Return Norg(J)
    End Function
    Public Function dPorg(J As Integer) As Double
    Return Porg(J)
    End Function
    Public Function dKorg(J As Integer) As Double
    Return Korg(J)
    End Function
    Public Function dQorga(J As Integer) As Double
    Return Qorga(J)
    End Function
    Public Function sTypeMorga() As Double
    Return TypeMorga
    End Function
    Public Function dCsurNRes() As Double
    Return CsurNRes
    End Function
End Class
