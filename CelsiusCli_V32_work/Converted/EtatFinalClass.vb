Option Strict Off
Option Explicit Off
Imports System
Public Class EtatFinalClass
    Public Stockfinal As Double 'VX, Valeur du stock hydrique utile au dernier jour de la simulation (mm)
    Dim QpaillisFinal As Double
    Public FinalSolhautON As Boolean
     Dim Stsurf_fin As Double
    Dim Strac_fin As Double
    Dim Stnonrac_fin As Double
    Dim Stprofond_fin As Double
    Dim Stocksol_fin As Double
    Dim StockMes_fin As Double
    Dim Stger_fin As Double
    Dim FminY1_fin As Double
    Dim StockN1_fin As Double
    Dim StockC1_fin As Double
    Dim NminReliquat As Double
    'VERSION 3 du 8-10 nov 2017'
    'Modif FA dec 2021
    Public Sub EcritEtatFinal(stock As Double, Qpaillis As Double, SolHaut As Boolean)
    'version sommaire pour version celsius 2
    Stockfinal = stock
    QpaillisFinal = Qpaillis
    FinalSolhautON = SolHaut
    End Sub
    Public Sub EcritEF2(Sol As SolClass, plante As PLanteClass, mulch As MulchClass, nbjsimul As Integer)
    'version complete (Celsius 3.2)

    Stsurf_fin = Sol.dStsurf(nbjsimul)
    Strac_fin = Sol.dStrac(nbjsimul)
    Stnonrac_fin = Sol.dStnonrac(nbjsimul)
    Stprofond_fin = Sol.dStprofond(nbjsimul)
    Stocksol_fin = Sol.dStockSol(nbjsimul)
    StockMes_fin = Sol.dStockMes(nbjsimul)
    Stger_fin = Sol.dStger(nbjsimul)
    FminY1_fin = Sol.dFminY1
    StockN1_fin = Sol.dStockN1
    StockC1_fin = Sol.dStockC1
    QpaillisFinal = mulch.dQpaillis(nbjsimul)
    If Sol.dNreliquat > 0 Then NminReliquat = Sol.dNreliquat Else NminReliquat = 0

    End Sub
    Public Function dStsurf_fin() As Double
    Return Stsurf_fin
    End Function
    Public Function dStrac_fin() As Double
    Return Strac_fin
    End Function
    Public Function dStnonrac_fin() As Double
    Return Stnonrac_fin
    End Function
    Public Function dStprofond_fin() As Double
    Return Stprofond_fin
    End Function
    Public Function dStocksol_fin() As Double
    Return Stocksol_fin
    End Function
    Public Function dStockMes_fin() As Double
    Return StockMes_fin
    End Function
    Public Function dStger_fin() As Double
    Return Stger_fin
    End Function
    Public Function dFminY1_fin() As Double
    Return FminY1_fin
    End Function
    Public Function dStockN1_fin() As Double
    Return StockN1_fin
    End Function
    Public Function dStockC1_fin() As Double
    Return StockC1_fin
    End Function
    Public Function dQpaillisFinal() As Double
    Return QpaillisFinal
    End Function
    Public Function dNminReliquat() As Double
    Return NminReliquat
    End Function
End Class
