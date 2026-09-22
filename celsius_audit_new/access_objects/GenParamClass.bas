Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Codgenclim As Boolean 'si utilisation d'un générateur de climat --a déplacer vers optionsmodel ?
Dim Vlaimax As Double 'PX constante d'influence du taux de développement sur le LAI
Dim ParSurRg As Double 'PX

Dim rstGenParam As ADODB.Recordset
'lecture des paramètres généraux de simulation
'table "general_paramaters"
'VERSION 3 du 8-10 nov 2017'

Sub LisGenParam(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)


Dim Trouve As Boolean
Set rstGenParam = New ADODB.Recordset
Trouve = False

rstGenParam.Open "General_Parameters", DataBase_Cnn
rstGenParam.MoveFirst
While Not rstGenParam.EOF And Not Trouve
    If rstGenParam!idGenParam = SimUnit.sidGenParam Then
        
        Trouve = True
        Vlaimax = rstGenParam!Vlaimax
        Codgenclim = rstGenParam!Codgenclim
        ParSurRg = rstGenParam!ParSurRg
                            
    End If
    rstGenParam.MoveNext
Wend
'fermeture table et libération mémoire de l'objet

rstGenParam.Close
Set rstGenParam = Nothing

End Sub

Property Get nCodgenclim() As Integer
nCodgenclim = Codgenclim
End Property
Property Get dVlaimax() As Double
dVlaimax = Vlaimax
End Property
Property Get dParSurRg() As Double
dParSurRg = ParSurRg
End Property