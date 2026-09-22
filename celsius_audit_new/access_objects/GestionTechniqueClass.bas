Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'VX: Variable explicative, VE: variable d'état, VS: variable simulée,PX: paramètre, CS: controle de simulation (lien vers autres variables)
Dim rstDataTech As ADODB.Recordset
Dim IdCultivar(1 To 2) As String 'champ identifiant l'espece et le cultivar, lu par le modèle et utilisé pour trouver les paramètres du cultivar et de l'espece dans les tables correspondantes
Dim Nbcult As Integer 'nombre de cultivars dans l'association, lu dans "Tech_Commun"
Dim Numcrop As Integer 'numéro du cultivar de l'association, lu dans "Tech_perCrop"
Dim TypeInstal(1 To 2) As Integer 'VX, Type d'installation du cultivar de numéro Numcrop. 1: semis graine pré-germée; 2: semis graine sèche; 3: repiquage
Dim RepiquageON(1 To 2) As Boolean 'VX, repiquage oui non (pour le cultivar numcrop)

Dim iplt(1 To 2) As Integer 'VX, date de semis en jour de 1 à 731 pour le cult. numcrop
'Dim jourplt(1 To 2) As Integer
' jourplt est dans le calendrier de simulation ***PAS UTILISE
Dim DensSem(1 To 2) As Single 'VX, densité de plantes au semis pour le cultivar numcrop
Dim irepiqu(1 To 2) As Integer 'VX jour de repiquage ou démariage pour le cultivar numcrop
Dim densrepiqu(1 To 2) As Single 'VX,densité après repiquage ou démariage pour le cultivar numcrop
Dim ilev(1 To 2) As Integer 'VX, jour de la levée dans le calendrieer annuel pour le cultivar numcrop
Dim SerreTunnelON As Boolean 'VX, Si vrai pépinière en serre tunnel, la température de l'air sera corrigée par procédure CorrigTSerres pendant tout la durée de la pépinière (cropsta=3)
Dim imulch As Integer 'VX date mise en place du mulch
Dim CodParamMulch As Integer 'CS code du type de mulch dans la table "mulch"
Dim QpaillisApport As Double 'VX quantité de paillis apportée au jour imulch (Mg/ha)
Dim AltiCult As Integer 'VX, altitude de la culture pour correction temperature station
Dim IrrigON As Boolean 'VX, culture irriguée Vrai/ Faux

Dim tApportMON As Double 'VX N apporte par amendements organiques (K.ha-1)
Dim tApportMinN As Double 'VX apport en N minéral (engrais minéral - peut contenir reliquats Nitrate et ammonium dans le sol au début de la simulation)
Dim CsurN_AO As Double 'VX  rapport C sur N de l'amendement organique moyen (bilan saisonnier de N)
Dim KresY As Double 'VX Taux moyen saisonnier de minéralisation de l'amendement organique (kg.ha-1.an-1)
'variables de la gestion technique auto
Dim DbutoirNouvSemis(1 To 2) As Integer 'VX: date en jours après levée au delà de laquelle une culture détruite par un aléa n'est plus ressemée
Dim Ressemis(1 To 2) As Boolean 'VE ressemis à faire vrai ou faux (si CyberST activé dans options model, devient vrai en cas de mort de la plante avant un certain jour
Dim SeuilCumPrecip(1 To 2) As Double 'VX seuil de cumul de pluies infiltrées consécutives déclenchant le semis en cas de semis automatique
Dim JourNouvSem(1 To 2) As Integer
Dim Nbsemis(1 To 2)
Dim fertiminON As Boolean '??
Dim fertiorgON As Boolean '??
Dim DriveRuiObs As Boolean 'VX, si vrai forçage du modèle par le ruissellement observé, lu dans table RuissellementObs
'
'VERSION 3 du 8-10 nov 2017'
' Modifs FA novembre 2021 bilan saisonnier N et C Hénin Dupuis
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026


Sub LisTech(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)
Dim Trouve As Boolean
Dim msg As String

Set rstDataTech = New ADODB.Recordset
Trouve = False
' lecture Tech_Commun

rstDataTech.Open "Tech_Commun", DataBase_Cnn
rstDataTech.MoveFirst
While Not rstDataTech.EOF And Not Trouve
    If rstDataTech!IdTech_Com = SimUnit.sIdTec Then
        Trouve = True
        Nbcult = rstDataTech!Nbcult
        SerreTunnelON = rstDataTech!SerreTunnelON
        imulch = rstDataTech!imulch
        CodParamMulch = rstDataTech!CodParamMulch
        QpaillisApport = rstDataTech!QpaillisApport
        AltiCult = rstDataTech!AltiCult
        IrrigON = rstDataTech!IrrigON
        tApportMON = rstDataTech!tApportMON
        tApportMinN = rstDataTech!tApportMinN
        fertiminON = rstDataTech!fertiminON
        fertiorgON = rstDataTech!fertiorgON
        CsurN_AO = rstDataTech!CsurN_AO
        KresY = rstDataTech!KresY
        DriveRuiObs = rstDataTech!DriveRuiObs
                
    End If
 rstDataTech.MoveNext
Wend
rstDataTech.Close
' vérifier si faut pas rstDataTech=nothing
Set rstDataTech = New ADODB.Recordset

'Reading Tech_perCrop

rstDataTech.Open "SELECT * FROM Tech_perCrop where IdTech_Com='" & SimUnit.sIdTec & "' Order by NumCrop", DataBase_Cnn
rstDataTech.MoveFirst
While Not rstDataTech.EOF
        Numcrop = rstDataTech!Numcrop
        IdCultivar(Numcrop) = rstDataTech!IdCultivar
        TypeInstal(Numcrop) = rstDataTech!TypInstal
        RepiquageON(Numcrop) = rstDataTech!RepiquageON
        iplt(Numcrop) = rstDataTech!isem
        DensSem(Numcrop) = rstDataTech!DensSem
        ilev(Numcrop) = rstDataTech!ilev
        irepiqu(Numcrop) = rstDataTech!irepiqu
        densrepiqu(Numcrop) = rstDataTech!densrepiqu
        DbutoirNouvSemis(Numcrop) = rstDataTech!DbutoirNouvSemis
        SeuilCumPrecip(Numcrop) = rstDataTech!SeuilCumPrecip
        Ressemis(Numcrop) = rstDataTech!SemisAutoDebut

        Call CoherenceGTech(Numcrop)
    rstDataTech.MoveNext
Wend
' test cohérence numcrop NbCult
If Numcrop <> Nbcult Then
    Nbcult = Numcrop
    msg = "nombre de cultures retenues: " & Str(Numcrop)
    MsgBox (msg)
End If
'fermeture table et libération mémoire de l'objet
rstDataTech.Close
Set rstDataTech = Nothing

End Sub
Sub CoherenceGTech(Numcrop As Integer)
Dim CoherenceSemis As Boolean
Dim CoherenceIlev As Boolean
Dim CoherenceDensite As Boolean
Dim msg As String

CoherenceSemis = True
CoherenceIlev = True
CoherenceDensite = True

If (TypeInstal(Numcrop) = 3 And irepiqu(Numcrop) <> 999) Then CoherenceSemis = False
If (TypeInstal(Numcrop) = 3 And irepiqu(Numcrop) <> 999) Then CoherenceSemis = False
If (TypeInstal(Numcrop) = 3 And Not RepiquageON(Numcrop) And irepiqu(Numcrop) <> 999) Then CoherenceSemis = False

If Not CoherenceSemis Then
    msg = "Attention pb potentiel de coherence date repiquage / type d'installation"
    msg = msg & "si pas de repiquage (TypInstal<>3) et RepiquageON=non, irepiqu doit être 999"
    MsgBox (msg)
    msg = "seule possibilité où typinstal=3 et repiquageON=Oui : il y a démariage à irepiqu (réduction de la densité)"
    MsgBox (msg)
End If
If TypeInstal(Numcrop) = 1 And ilev(Numcrop) <> iplt(Numcrop) + 1 Then CoherenceIlev = False

If Not CoherenceIlev Then
    msg = "pb de coherence date de levée, typinstal, date semis"
    msg = msg & "si semis graine prégermée, typinstal=1 et datelevée doit être = isem+1"
    MsgBox (msg)
End If
If RepiquageON(Numcrop) = False And densrepiqu(Numcrop) <> DensSem(Numcrop) Then CoherenceDensite = False
If Not CoherenceDensite Then
    msg = " pb de coherence densité de repiquage / densité de semis "
    MsgBox (msg)
End If
End Sub
Sub CyberPlouck(icult As Integer, Die As Boolean, Deathday As Integer, Joursim As Integer, StartDay As Integer)
'automate ajustant des actes techniques aux conditions du milieu
If (Die And (Deathday > 0)) And (Joursim + StartDay - 1 < DbutoirNouvSemis(icult)) Then
    Ressemis(icult) = True
    
End If
End Sub
Sub SemisAuto(icult As Integer, Joursim As Integer, StartDay As Integer, Cumpluinf As Double)
Dim nouvsemJA As Integer, EcartSem As Integer

If (Joursim + StartDay - 1 >= iplt(icult)) And (Cumpluinf >= SeuilCumPrecip(icult)) Then
    JourNouvSem(icult) = Joursim
    nouvsemJA = Joursim + StartDay - 1
    EcartSem = nouvsemJA - iplt(icult)
    If irepiqu(icult) <> 999 Then irepiqu(icult) = irepiqu(icult) + EcartSem
    iplt(icult) = nouvsemJA
    'histoire de ne plus repasser par ici:
    Ressemis(icult) = False
    '
    ' attention il y a peut etre besoin de réinitialiser la plante...?
End If
End Sub
Sub CompteSemis(icult As Integer)
Nbsemis(icult) = Nbsemis(icult) + 1
End Sub
Property Get sIdCultivar(icult As Integer) As String
sIdCultivar = IdCultivar(icult)
End Property
Property Get nNbCult() As Integer
nNbCult = Nbcult
End Property
Property Get niplt(icult As Integer) As Integer
niplt = iplt(icult)
End Property
Property Get nilev(icult As Integer) As Integer
nilev = ilev(icult)
End Property
Property Get nirepiqu(icult As Integer) As Integer
nirepiqu = irepiqu(icult)
End Property
Property Get sTypeInstal(icult As Integer) As Integer
sTypeInstal = TypeInstal(icult)
End Property
Property Get bRepiquageON(icult As Integer) As Boolean
bRepiquageON = RepiquageON(icult)
End Property
Property Get dDensSem(icult As Integer) As Double
dDensSem = DensSem(icult)
End Property
Property Get dDensRepiqu(icult As Integer) As Double
dDensRepiqu = densrepiqu(icult)
End Property
Property Get bSerreTunnelON() As Boolean
bSerreTunnelON = SerreTunnelON
End Property
Property Get nCodParamMulch() As Integer
nCodParamMulch = CodParamMulch
End Property
Property Get nimulch() As Integer
nimulch = imulch
End Property
Property Get dQpaillisApport() As Double
dQpaillisApport = QpaillisApport
End Property
Property Get nAltiCult() As Integer
nAltiCult = AltiCult
End Property
Property Get bIrrigON() As Boolean
bIrrigON = IrrigON
End Property
Property Get bfertiminON() As Boolean
bfertiminON = fertiminON
End Property
Property Get bfertiorgON() As Boolean
bfertiorgON = fertiorgON
End Property
Property Get dtApportMON() As Double
dtApportMON = tApportMON
End Property
Property Get bRessemis(icult As Integer) As Boolean
bRessemis = Ressemis(icult)
End Property
Property Get nJourNouvSem(icult As Integer) As Integer
nJourNouvSem = JourNouvSem(icult)
End Property
Property Get nNbsemis(icult As Integer) As Integer
nNbsemis = Nbsemis(icult)
End Property
Property Get dtApportMinN() As Double
dtApportMinN = tApportMinN
End Property
Property Get bDriveRuiObs() As Boolean
bDriveRuiObs = DriveRuiObs
End Property
Property Get dCsurN_AO() As Double
dCsurN_AO = CsurN_AO
End Property
Property Get dKresY() As Double
dKresY = KresY
End Property