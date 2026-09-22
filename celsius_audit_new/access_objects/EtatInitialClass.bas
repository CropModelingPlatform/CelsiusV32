Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Stockinit As Double 'VX, Valeur du stock hydrique utile au premier jour de la simulation (mm)
Dim Qpaillisinit As Double
Dim IniSolhautON As Boolean 'VX, Choix d'initialiser le stock par le haut (Oui) ou par le bas du profil (Non)
Dim rstDataIni As ADODB.Recordset
Dim Ninit As Double
'VERSION 3 du 8-10 nov 2017'
' modif FA du 18/12/21 ajout Ninit

Sub LisInitialData(DataBase_Cnn As ADODB.Connection, idIni As String)


Dim Trouve As Boolean
Set rstDataIni = New ADODB.Recordset
Trouve = False

rstDataIni.Open "ParamIni", DataBase_Cnn
rstDataIni.MoveFirst
While Not rstDataIni.EOF And Not Trouve
    If rstDataIni!idIni = idIni Then
        
        Trouve = True
        Stockinit = rstDataIni!Stockinit
        Qpaillisinit = rstDataIni!Qpaillisinit
        IniSolhautON = rstDataIni!IniSolhautON
        Ninit = rstDataIni!Ninit
    End If
    rstDataIni.MoveNext
Wend
'fermeture table et libération mémoire de l'objet

rstDataIni.Close
Set rstDataIni = Nothing

End Sub
Sub recursive(EtatFinal As EtatFinalClass)
Stockinit = EtatFinal.Stockfinal
        Qpaillisinit = EtatFinal.dQpaillisFinal
        IniSolhautON = EtatFinal.FinalSolhautON
        Ninit = EtatFinal.dNminReliquat
End Sub
Property Get dStockinit() As Double
dStockinit = Stockinit
End Property
Property Get dQpaillisinit() As Double
dQpaillisinit = Qpaillisinit
End Property
Property Get bIniSolhautON() As Boolean
bIniSolhautON = IniSolhautON
End Property
Property Get dNinit() As Double
dNinit = Ninit
End Property