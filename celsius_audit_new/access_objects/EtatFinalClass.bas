Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

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

Sub EcritEtatFinal(stock As Double, Qpaillis As Double, SolHaut As Boolean)
'version sommaire pour version celsius 2
Stockfinal = stock
QpaillisFinal = Qpaillis
FinalSolhautON = SolHaut
End Sub


Sub EcritEF2(Sol As SolClass, plante As PLanteClass, mulch As MulchClass, nbjsimul As Integer)
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
Property Get dStsurf_fin() As Double
dStsurf_fin = Stsurf_fin
End Property
Property Get dStrac_fin() As Double
dStrac_fin = Strac_fin
End Property
Property Get dStnonrac_fin() As Double
dStnonrac_fin = Stnonrac_fin
End Property
Property Get dStprofond_fin() As Double
dStprofond_fin = Stprofond_fin
End Property
Property Get dStocksol_fin() As Double
dStocksol_fin = Stocksol_fin
End Property
Property Get dStockMes_fin() As Double
dStockMes_fin = StockMes_fin
End Property
Property Get dStger_fin() As Double
dStger_fin = Stger_fin
End Property
Property Get dFminY1_fin() As Double
dFminY1_fin = FminY1_fin
End Property
Property Get dStockN1_fin() As Double
dStockN1_fin = StockN1_fin
End Property
Property Get dStockC1_fin() As Double
dStockC1_fin = StockC1_fin
End Property
Property Get dQpaillisFinal() As Double
dQpaillisFinal = QpaillisFinal
End Property
Property Get dNminReliquat() As Double
dNminReliquat = NminReliquat
End Property