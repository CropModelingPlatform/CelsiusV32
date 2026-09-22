Option Strict Off
Option Explicit Off
Imports System
Public Class EtatInitialClass
    Dim Stockinit As Double 'VX, Valeur du stock hydrique utile au premier jour de la simulation (mm)
    Dim Qpaillisinit As Double
    Dim IniSolhautON As Boolean 'VX, Choix d'initialiser le stock par le haut (Oui) ou par le bas du profil (Non)
    Dim rstDataIni As ADODB.Recordset
    Dim Ninit As Double
    'VERSION 3 du 8-10 nov 2017'
    ' modif FA du 18/12/21 ajout Ninit
    Public Sub LisInitialData(DataBase_Cnn As ADODB.Connection, idIni As String)


    Dim Trouve As Boolean
    rstDataIni = New ADODB.Recordset
    Trouve = False

    rstDataIni.Open("ParamIni", DataBase_Cnn)
    rstDataIni.MoveFirst
    While Not rstDataIni.EOF And Not Trouve
        If rstDataIni("idIni") = idIni Then

            Trouve = True
            Stockinit = rstDataIni("Stockinit")
            Qpaillisinit = rstDataIni("Qpaillisinit")
            IniSolhautON = rstDataIni("IniSolhautON")
            Ninit = rstDataIni("Ninit")
        End If
        rstDataIni.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet

    rstDataIni.Close
    rstDataIni = Nothing

    End Sub
    Public Sub recursive(EtatFinal As EtatFinalClass)
    Stockinit = EtatFinal.Stockfinal
            Qpaillisinit = EtatFinal.dQpaillisFinal
            IniSolhautON = EtatFinal.FinalSolhautON
            Ninit = EtatFinal.dNminReliquat
    End Sub
    Public Function dStockinit() As Double
    Return Stockinit
    End Function
    Public Function dQpaillisinit() As Double
    Return Qpaillisinit
    End Function
    Public Function bIniSolhautON() As Boolean
    Return IniSolhautON
    End Function
    Public Function dNinit() As Double
    Return Ninit
    End Function
End Class
