Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim rstSolData As ADODB.Recordset 'les tables de variables sol lues
Dim CdeRui As Integer 'copie locale de Typerui, champ de laison avec la table sol
Dim Ap1 As Double 'Vx parametre Ap1 des fonctions de ruissellement Albergel et al, et coef de ruissellement (si Ap2..Ap4=0, Rui=Ap1*(P-Seuil))
Dim Ap2 As Double, Ap3 As Double, Ap4 As Double 'Vx parametre Ap2 des fonctions de ruissellement Albergel et al combinée avec effet paillis (Ruis(Joursim) = (Ap1 + Ap3 * IKJ + b_ruis * Qpaillis(Joursim)) * (precip - (Ap4 - Ap2 * IKJ) / (Ap1 + Ap3 * IKJ))
Dim seuil_ruis As Double 'VX seuil de précipitations en-dessous duquel il n'y a pas de ruissellement (mm). Non utilisé si Ap2 ou Ap3 ou Ap4 <> 0
Dim EffetLAI As Boolean 'VX: vrai: effet du LAI sur le ruissellement
'   lecture des paramètres d'état de surface du sol pour simulation ruissellement (dans la classe MulchClass)
'VERSION 3 du 8-10 nov 2017
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026

Sub readSurf(DataBase_Cnn As ADODB.Connection, TypeRui As Integer)
Set rstSolData = New ADODB.Recordset

rstSolData.Open "SELECT * FROM TypeSurfSol WHERE TypeRui = " & TypeRui, DataBase_Cnn
CdeRui = TypeRui
Ap1 = rstSolData!Ap1
Ap2 = rstSolData!Ap2
Ap3 = rstSolData!Ap3
Ap4 = rstSolData!Ap4
seuil_ruis = rstSolData!seuil_ruis
EffetLAI = rstSolData!EffetLAI
rstSolData.Close
Set rstSolData = Nothing
End Sub
Property Get dAp1() As Double
dAp1 = Ap1
End Property
Property Get dAp2() As Double
dAp2 = Ap2
End Property
Property Get dAp3() As Double
dAp3 = Ap3
End Property
Property Get dAp4() As Double
dAp4 = Ap4
End Property
Property Get dseuil_ruis() As Double
dseuil_ruis = seuil_ruis
End Property
Property Get bEffetLAI() As Boolean
bEffetLAI = EffetLAI
End Property
Property Get nCdeRui() As Integer
nCdeRui = CdeRui
End Property