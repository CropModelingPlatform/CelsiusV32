Option Strict Off
Option Explicit Off
Imports System
Public Class GenParamClass
    Dim Codgenclim As Boolean 'si utilisation d'un générateur de climat --a déplacer vers optionsmodel ?
    Dim Vlaimax As Double 'PX constante d'influence du taux de développement sur le LAI
    Dim ParSurRg As Double 'PX

    Dim rstGenParam As ADODB.Recordset
    'lecture des paramètres généraux de simulation
    'table "general_paramaters"
    'VERSION 3 du 8-10 nov 2017'
    Public Sub LisGenParam(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)


    Dim Trouve As Boolean
    rstGenParam = New ADODB.Recordset
    Trouve = False

    rstGenParam.Open("General_Parameters", DataBase_Cnn)
    rstGenParam.MoveFirst
    While Not rstGenParam.EOF And Not Trouve
        If rstGenParam("idGenParam") = SimUnit.sidGenParam Then

            Trouve = True
            Vlaimax = rstGenParam("Vlaimax")
            Codgenclim = rstGenParam("Codgenclim")
            ParSurRg = rstGenParam("ParSurRg")

        End If
        rstGenParam.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet

    rstGenParam.Close
    rstGenParam = Nothing

    End Sub

    Public Function nCodgenclim() As Integer
    Return Codgenclim
    End Function
    Public Function dVlaimax() As Double
    Return Vlaimax
    End Function
    Public Function dParSurRg() As Double
    Return ParSurRg
    End Function
End Class
