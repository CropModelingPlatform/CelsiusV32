Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Qpaillis(0 To 731) As Double 'VE quantité de paillis présente au sol chaque jour (Mg/ha)
Dim Eomulch As Double 'VE, évaporation potentielle du mulch (mm)
Dim Emulch(0 To 731) As Double 'VE, évaporation du mulch (mm)
Dim Smulch As Double 'VE, stock en eau du mulch (mm)
Dim EauVersSol As Double 'VE, quantité d'eau disponible sous le mulch pour infiltration dans le sol (mm)
Dim Eau_vers_Mulch As Double 'VE quantité d'eau disponible après ruissellement pour infiltration dans le sol et le mulch(mm)
Dim rstMulchData As ADODB.Recordset 'table des données des caractéristiques du mulch
Dim gamma_mulch As Double 'VX, Coefficient d'extinction de l'évapotranspiration potentielle par le mulch
Dim CapaciteWMulch As Double 'VX, capacité de stockage de l'eau par le mulch en mm/T/ha de mulch
Dim Alpha_pail As Double 'VX, exp(-Alph_pail) est le taux de disparition quotidienne du paillis
Dim FracSoilCover As Double 'VE fraction du sol couverte par le paillis (sd)
Dim Beta_pail As Double 'VX, pouvoir couvrant du paillis (ha/T DM)
Dim b_ruis As Double    'VX, coefficient d'augmentation du ruissellement par le paillis (est en général négatif car le paillis réduit le ruissellement)
Dim Ruis(0 To 731) As Double 'VE, ruissellement journalier (mm)
Dim SigmaSimEmulch As Double 'VS cumul sur la simulation de Emulch (mm)
Dim SigmaSimRuis As Double 'VS cumul sur la simulation de Ruis (mm)
Dim SigmaSimPluM As Double 'VS, cumul sur la simulation de la pluie parvenant à la couche de mulch (avant ruissellement en surface) (mm)
Dim IKJ As Double 'VE indice d'antériorité des pluies (Albergel et al.) pour le calcul du ruissellement au sahel
Dim RuisEtrange As Boolean
Dim rstDataRuiObs As ADODB.Recordset
'VERSION 3 du 8-10 nov 2017'
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026


Sub LisMulch(DataBase_Cnn As ADODB.Connection, CodParamMulch As Integer, Qpaillisinit As Double)
Dim Trouve As Boolean
Dim msg As String
Set rstMulchData = New ADODB.Recordset

Trouve = False

rstMulchData.Open "Mulch", DataBase_Cnn
rstMulchData.MoveFirst
While Not rstMulchData.EOF And Not Trouve
    If rstMulchData!idMulch = CodParamMulch Then
        
        Trouve = True
        
        gamma_mulch = rstMulchData!gamma_mulch
        CapaciteWMulch = rstMulchData!CapaciteWMulch
        Alpha_pail = rstMulchData!Alpha_pail
        Beta_pail = rstMulchData!Beta_pail
        b_ruis = rstMulchData!b_ruis
        
    End If
    rstMulchData.MoveNext
Wend
rstMulchData.Close
Set rstMulchData = Nothing
Qpaillis(0) = Qpaillisinit
If Qpaillisinit > 0 And Alpha_pail = 0 Then
    msg = "Attention! Paillis initial non nul dans ParamIni et pas de décomposition du paillage (alpha_pail=0 dans Mulch) ! Le paillis sera initialisé à 0"
    MsgBox (msg)
    Qpaillis(0) = 0
End If
End Sub
Sub LisRuiObs(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)

Dim AnRuiObs As Integer
Dim JourRuiObs As Integer
Dim RuiObs As Double
Dim Joursim As Integer
Dim ErrFa As Boolean
Dim jourEnPlus As Integer
Dim Ndyear1 As Integer
Dim msg As String
Dim NbreRuiObs As Integer
On Error GoTo Err_LisRuiObs

If SimUnit.bSY_Bissextile Then Ndyear1 = 366 Else Ndyear1 = 365
Set rstDataRuiObs = New ADODB.Recordset
'Lecture table "RuissellementObs" à modifier si nécessaire
rstDataRuiObs.Open "SELECT * FROM RuissellementObs where idTechCom='" & SimUnit.sIdTec & "' Order by YearRuiObs, JourRuiObs", DataBase_Cnn, adOpenDynamic
rstDataRuiObs.MoveFirst


While Not rstDataRuiObs.EOF
    

    AnRuiObs = rstDataRuiObs!YearRuiObs

    JourRuiObs = rstDataRuiObs!JourRuiObs
    
    RuiObs = rstDataRuiObs!RuiObs
    
    If AnRuiObs = SimUnit.nStartYear Then
        Joursim = JourRuiObs - SimUnit.nStartDay + 1
    Else
        Joursim = JourRuiObs + Ndyear1 - SimUnit.nStartDay + 1
    End If
    If Joursim < 0 Then
        msg = "Attention, calendrier des irrigations incohérent avec calendrier de simulation"
        GoTo ErrFA_LisRuiObs
    Else
        Ruis(Joursim) = RuiObs
        NbreRuiObs = NbreRuiObs + 1
    End If
    rstDataRuiObs.MoveNext
Wend



If NbreRuiObs = 0 Then
    msg = "error in observed runoff table, no runoff found during simulation period"
    GoTo ErrFA_LisRuiObs
End If

rstDataRuiObs.Close
Set rstDataRuiObs = Nothing



Exit_LisRuiObs:
    Exit Sub
ErrFA_LisRuiObs:
    ErrFa = True
    MsgBox ("ERREUR ! " & msg)
    Resume Exit_LisRuiObs
Err_LisRuiObs:
    MsgBox Err.Description
    Resume Exit_LisRuiObs

End Sub
Sub BiomasseMulch(Joursim As Integer, jourpaillage As Boolean, QpaillisApport As Double)
Qpaillis(Joursim) = Qpaillis(Joursim - 1) * Exp(-Alpha_pail)
If jourpaillage Then Qpaillis(Joursim) = Qpaillis(Joursim) + QpaillisApport
FracSoilCover = 1 - Exp(-Beta_pail * Qpaillis(Joursim))
End Sub
Sub BilanMulch(Joursim As Integer, EoSM As Double, precip As Double)
'calcul du bilan hydrique du mulch équivalent à la routine introduite dans Stics 3.0 par A. Findeling
' selon Arreola 1996
'precip est l'eau de pluie moins le ruissellement

Dim epail1 As Double
Dim epail2 As Double
Dim intercep As Double
Dim msg As String

Eomulch = EoSM * (1 - Exp(-gamma_mulch * Qpaillis(Joursim)))
    If Qpaillis(Joursim) > 0 And Qpaillis(Joursim - 1) > 0 Then
        
         Smulch = Smulch * Qpaillis(Joursim) / Qpaillis(Joursim - 1)
         If Smulch > 0 Then
'  calcul du premier terme de l'évaporation du paillis, epail1, dû à la
'  disparition d'une quantité (Qpaillis(joursim-1) - Qpaillis(joursim) et donc de l'eau
' qu 'elle contenait
           epail1 = Smulch * ((Qpaillis(Joursim - 1) - Qpaillis(Joursim)) / Qpaillis(Joursim))
           
           If epail1 < 0 Then epail1 = 0
           If epail1 < Eomulch Then
'  calcul du deuxième terme d'évaporation calculé comme le complément de ep1
'  à l'évap potentielle eopaillis en respectant la contrainte de stoc : respail>0
               epail2 = Eomulch - epail1
               If epail2 > Smulch Then epail2 = Smulch
           End If
         Else
             Smulch = 0
         End If
    
    End If

Emulch(Joursim) = epail1 + epail2
If epail1 + epail2 > Smulch Then
msg = "yabug"
End If
' on passe par FracSoilCover pour l'interception de l'eau car le parametre capaciteWmulch
' peut ainsi être mesuré par gravimétrie sur un échantillon de pailles et qu'on peut empiriquement
' déduire aussi la relation entre quantité de mulch et taux de couverture du sol
' mais attention, dans stics6 il y a un contresens, FracSoilCover n'est pas
' utilisé pour le stockage de l'eau mais il l'est pour l'évaporation (selon le bouquin en tout cas)!!

intercep = precip * FracSoilCover
Smulch = Smulch - epail2 + intercep
If Smulch > CapaciteWMulch * Qpaillis(Joursim) Then

    EauVersSol = precip - intercep + Smulch - CapaciteWMulch * Qpaillis(Joursim)
    Smulch = CapaciteWMulch * Qpaillis(Joursim)

Else
    EauVersSol = precip - intercep
'à ce stade Smulch contient encore toute l'eau de pluie du jour et on lui enlève l'excès ligne suivante
'si pas de mulch Smulch est égal à précip, Eauversol aussi et Smulch devient nul
End If

SigmaSimEmulch = SigmaSimEmulch + Emulch(Joursim)

End Sub
Sub Ruissellement(Joursim As Integer, precip As Double, TypeSurf As TypeSurfClass, Lai As Double)
'selon  albergel et al. mixé avec effet mulch Scopel et al (comme proposé par B. Rapidel, thèse de Fagaye Sissoko)
Dim seuil As Double

SigmaSimPluM = SigmaSimPluM + precip

If Ruis(Joursim) = 0 And (precip > 0) Then
' si pas de ruissellement observé

    If TypeSurf.dAp2 = 0 And TypeSurf.dAp3 = 0 And TypeSurf.dAp4 = 0 Then
        seuil = TypeSurf.dseuil_ruis
    Else
        If TypeSurf.dAp1 + TypeSurf.dAp3 * IKJ = 0 Then
            seuil = 0
        Else
            seuil = (TypeSurf.dAp4 - TypeSurf.dAp2 * IKJ) / (TypeSurf.dAp1 + TypeSurf.dAp3 * IKJ)
        End If
    End If

    Ruis(Joursim) = (TypeSurf.dAp1 + TypeSurf.dAp3 * IKJ + b_ruis * Qpaillis(Joursim)) * (precip - seuil)
    Ruis(Joursim) = Max(0, Ruis(Joursim))
    If TypeSurf.bEffetLAI Then Ruis(Joursim) = Ruis(Joursim) * Exp(-0.5 * Lai)
Else
    'cas d'un ruissellement observé non nul alors que pluie+irrig=0, on force à 0 le ruissellement mais on le signale
    If Ruis(Joursim) <> 0 And (precip = 0) Then
        Ruis(Joursim) = 0
        RuisEtrange = True
    End If
   ' dans le cas d'un ruissellement observé, c'est sa valeur qui est retenue et utilisée plus loin
End If
' mise à jour indice antériorité des pluies IKJ pour la prochaine itération: (IKJ(n+1)=(IKJ(n)+P(n))*exp(-0.5)
IKJ = (IKJ + precip) * Exp(-0.5)

SigmaSimRuis = SigmaSimRuis + Ruis(Joursim)
Eau_vers_Mulch = precip - Ruis(Joursim)

End Sub
Property Get dEomulch() As Double
dEomulch = Eomulch
End Property
Property Get dEmulch(Joursim As Integer) As Double
dEmulch = Emulch(Joursim)
End Property
Property Get dEauVersSol() As Double
dEauVersSol = EauVersSol
End Property
Property Get dEau_vers_Mulch() As Double
dEau_vers_Mulch = Eau_vers_Mulch
End Property
Property Get dSigmaSimEmulch() As Double
dSigmaSimEmulch = SigmaSimEmulch
End Property
Property Get dSigmaSimRuis() As Double
dSigmaSimRuis = SigmaSimRuis
End Property
Property Get dRuis(Joursim As Integer) As Double
dRuis = Ruis(Joursim)
End Property
Property Get dQpaillis(Joursim As Integer) As Double
dQpaillis = Qpaillis(Joursim)
End Property
Property Get dSigmaSimPluM() As Double
dSigmaSimPluM = SigmaSimPluM
End Property
Property Get bRuisEtrange() As Boolean
bRuisEtrange = RuisEtrange
End Property