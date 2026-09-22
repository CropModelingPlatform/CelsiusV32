Option Compare Database
Option Explicit
Public Codesuite As Integer 'introduit pour simulations récursives:
'if codesuite  = 0, initial state taken in ParamIni. If value <>0, initial state = final state of the previous SimUnit in the table as ordered by ChampTri ascending FInal state is  stored in memory (not in a table)
 

Public Sub Principal()
'Module maître de CELSIUS (CEreal an Legume SImulator Under Savanah environment)
'VERSION 3 du 8-10 nov 2017'
' modif FA du 18/12/21 reglages de la recursivité
' modifs FA mai 2026 nettoyage codes non utilisés (dont version antérieure de la récursivité)

Dim Db_Cnn As ADODB.Connection

Dim SimInfo As New ADODB.Recordset
Dim ReqSim As New ADODB.Recordset
Dim TabSynt As New ADODB.Recordset
Dim NumSimul As Long
Dim CodeOptim As Boolean
Dim comptesim As Long
Dim MsgFin As String
Dim chronoStart As Date
Dim StartChronoCalc As Date
Dim StartChronoEcrit As Date
Dim Chrono1 As Date
Dim Chrono2 As Date
Dim chronotot As Date
Dim SimCtrl As New SimulationControlClass


chronoStart = Now

Set Db_Cnn = CurrentProject.Connection
' vidage tables sorties
TabSynt.Open "OutputSynt", Db_Cnn, , adLockOptimistic
While Not TabSynt.EOF
     TabSynt.Delete
     TabSynt.MoveNext
Wend



SimInfo.Open "SELECT Count(SimUnitList.idsim) AS nimul FROM SimUnitList", Db_Cnn
NumSimul = SimInfo.Fields(0)


comptesim = 0
Set SimInfo = Nothing


ReqSim.Open "SELECT * FROM SimUnitList Order by ChampTri", Db_Cnn
' ChampTri is a field in SimUnitList introduced to allow recursive simulations. Champtri should be numbering the simulations to be run in recursive sequence,
' with the field Codesuite of table SimUnitlist having to be set at 0 for the first SimUnit ii the sequence and 1 for all the other SimUnit in the sequence

' boucle sur la liste des simulations
While Not ReqSim.EOF
'extension possible : introduire ici boucle sur les années de simulation pour cas où EndYear-StartYear > 1 (manioc ou canne par exemple ??)
'prévoir de passer l'année en argument des appels pour controle des lectures de parametres en fonction de l'année...
'

    StartChronoEcrit = Now
    Forms("MenuPrincipal").Caption = "simulation numero " & comptesim + 1 & "/ " & NumSimul & "  idsim=" & SimCtrl.sIdentSimul
    DoEvents
 'piege à bugs specifique d'une sim unit
 'If comptesim + 1 = 135 Then
' NumSimul = NumSimul
' End If
'    If SimCtrl.sIdentSimul = "AgmipLT.ICGA_1973_2.C0N0" Then
'    NumSimul = NumSimul
'    End If
    Call SimCtrl.ReadParameters(ReqSim, Db_Cnn, comptesim)
    Chrono1 = Chrono1 + (Now - StartChronoEcrit)
    StartChronoCalc = Now
    Call SimCtrl.Simulation
    Chrono2 = Chrono2 + (Now - StartChronoCalc)
    StartChronoEcrit = Now
    Call SimCtrl.SortieSynthesis(TabSynt)
    Call SimCtrl.EcritDresu(Db_Cnn, comptesim)
    Call SimCtrl.MemoEtatFinal
    Chrono1 = Chrono1 + (Now - StartChronoEcrit)


'fin boucle possible sur les années de simulation si EndYear-StartYear > 1
ReqSim.MoveNext
comptesim = comptesim + 1

Wend
ReqSim.Close

Set ReqSim = Nothing
Set TabSynt = Nothing
'calculation and display of computing time :
chronotot = Now - chronoStart
MsgFin = "OK simulation complete; tps total= " & chronotot & "tps ecriture/lecture= " & Chrono1 & " tps calcul= " & Chrono2

If comptesim <> NumSimul Then MsgFin = "Attention ! Nombre de simulations différent du nombre indiqué"

MsgBox (MsgFin)

End Sub

Function Inverse(x) As Double
Dim msg As String
Dim ErrFa As Boolean

If x = 0 Then
    msg = "constantes thermiques nulles !"
    GoTo ErrFA_Inverse
End If
Inverse = 1 / x

Exit_Inverse:
    Exit Function
ErrFA_Inverse:
    ErrFa = True
    MsgBox ("ERREUR ! " & msg)
    Resume Exit_Inverse

End Function
Function Max(x, y) As Variant
If x > y Then Max = x Else Max = y
End Function
Function Min(x, y) As Variant
If x > y Then Min = y Else Min = x
End Function
Function Calendrier_versSim(jouraconvertir As Integer, jourdebut As Integer, Bissextile As Boolean)
' ********NON UTILISE introduit dans l'idée (pas mise en oeuvre donc) d'une
' utilisation à la place du tableau DOY(joursimul) qui permet de passer à tout moment du
' calendrier standard à celui de la simulation

' si le jouraconvertir est supérieur à 365, il sera interprété correctement dans le
'calendrier de simulation
' si le jouraconvertir est inférieur à jourdebut on considère qu'il concerne l'année suivant la 1ere année de simul
If jouraconvertir < jourdebut Then jouraconvertir = jouraconvertir + 365 - CInt(Bissextile)
Calendrier_versSim = jouraconvertir - jourdebut + 1
End Function
Function Zyva() As Boolean
'appel de CELSIUS par fonction pour bouton / macro
Call Principal
Zyva = True
End Function