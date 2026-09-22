Option Strict Off
Option Explicit Off
Imports System
Public Class GestionTechniqueClass
    'VX: Variable explicative, VE: variable d'état, VS: variable simulée,PX: paramètre, CS: controle de simulation (lien vers autres variables)
    Dim rstDataTech As ADODB.Recordset
    Dim IdCultivar(2) As String 'champ identifiant l'espece et le cultivar, lu par le modèle et utilisé pour trouver les paramètres du cultivar et de l'espece dans les tables correspondantes
    Dim Nbcult As Integer 'nombre de cultivars dans l'association, lu dans "Tech_Commun"
    Dim Numcrop As Integer 'numéro du cultivar de l'association, lu dans "Tech_perCrop"
    Dim TypeInstal(2) As Integer 'VX, Type d'installation du cultivar de numéro Numcrop. 1: semis graine pré-germée; 2: semis graine sèche; 3: repiquage
    Dim RepiquageON(2) As Boolean 'VX, repiquage oui non (pour le cultivar numcrop)

    Dim iplt(2) As Integer 'VX, date de semis en jour de 1 à 731 pour le cult. numcrop
    'Dim jourplt(1 To 2) As Integer
    ' jourplt est dans le calendrier de simulation ***PAS UTILISE
    Dim DensSem(2) As Single 'VX, densité de plantes au semis pour le cultivar numcrop
    Dim irepiqu(2) As Integer 'VX jour de repiquage ou démariage pour le cultivar numcrop
    Dim densrepiqu(2) As Single 'VX,densité après repiquage ou démariage pour le cultivar numcrop
    Dim ilev(2) As Integer 'VX, jour de la levée dans le calendrieer annuel pour le cultivar numcrop
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
    Dim DbutoirNouvSemis(2) As Integer 'VX: date en jours après levée au delà de laquelle une culture détruite par un aléa n'est plus ressemée
    Dim Ressemis(2) As Boolean 'VE ressemis à faire vrai ou faux (si CyberST activé dans options model, devient vrai en cas de mort de la plante avant un certain jour
    Dim SeuilCumPrecip(2) As Double 'VX seuil de cumul de pluies infiltrées consécutives déclenchant le semis en cas de semis automatique
    Dim JourNouvSem(2) As Integer
    Dim Nbsemis(2)
    Dim fertiminON As Boolean '??
    Dim fertiorgON As Boolean '??
    Dim DriveRuiObs As Boolean 'VX, si vrai forçage du modèle par le ruissellement observé, lu dans table RuissellementObs
    '
    'VERSION 3 du 8-10 nov 2017'
    ' Modifs FA novembre 2021 bilan saisonnier N et C Hénin Dupuis
    ' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026
    Public Sub LisTech(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass)
    Dim Trouve As Boolean
    Dim msg As String

    rstDataTech = New ADODB.Recordset
    Trouve = False
    ' lecture Tech_Commun

    rstDataTech.Open("Tech_Commun", DataBase_Cnn)
    rstDataTech.MoveFirst
    While Not rstDataTech.EOF And Not Trouve
        If rstDataTech("IdTech_Com") = SimUnit.sIdTec Then
            Trouve = True
            Nbcult = rstDataTech("Nbcult")
            SerreTunnelON = rstDataTech("SerreTunnelON")
            imulch = rstDataTech("imulch")
            CodParamMulch = rstDataTech("CodParamMulch")
            QpaillisApport = rstDataTech("QpaillisApport")
            AltiCult = rstDataTech("AltiCult")
            IrrigON = rstDataTech("IrrigON")
            tApportMON = rstDataTech("tApportMON")
            tApportMinN = rstDataTech("tApportMinN")
            fertiminON = rstDataTech("fertiminON")
            fertiorgON = rstDataTech("fertiorgON")
            CsurN_AO = rstDataTech("CsurN_AO")
            KresY = rstDataTech("KresY")
            DriveRuiObs = rstDataTech("DriveRuiObs")

        End If
     rstDataTech.MoveNext
    End While
    rstDataTech.Close
    ' vérifier si faut pas rstDataTech=nothing
    rstDataTech = New ADODB.Recordset

    'Reading Tech_perCrop

    rstDataTech.Open("SELECT * FROM Tech_perCrop where IdTech_Com='" & SimUnit.sIdTec & "' Order by NumCrop", DataBase_Cnn)
    rstDataTech.MoveFirst
    While Not rstDataTech.EOF
            Numcrop = rstDataTech("Numcrop")
            IdCultivar(Numcrop) = rstDataTech("IdCultivar")
            TypeInstal(Numcrop) = rstDataTech("TypInstal")
            RepiquageON(Numcrop) = rstDataTech("RepiquageON")
            iplt(Numcrop) = rstDataTech("isem")
            DensSem(Numcrop) = rstDataTech("DensSem")
            ilev(Numcrop) = rstDataTech("ilev")
            irepiqu(Numcrop) = rstDataTech("irepiqu")
            densrepiqu(Numcrop) = rstDataTech("densrepiqu")
            DbutoirNouvSemis(Numcrop) = rstDataTech("DbutoirNouvSemis")
            SeuilCumPrecip(Numcrop) = rstDataTech("SeuilCumPrecip")
            Ressemis(Numcrop) = rstDataTech("SemisAutoDebut")

            Call CoherenceGTech(Numcrop)
        rstDataTech.MoveNext
    End While
    ' test cohérence numcrop NbCult
    If Numcrop <> Nbcult Then
        Nbcult = Numcrop
        msg = "nombre de cultures retenues: " & Str(Numcrop)
    Console.Error.WriteLine((msg))
    End If
    'fermeture table et libération mémoire de l'objet
    rstDataTech.Close
    rstDataTech = Nothing

    End Sub
    Public Sub CoherenceGTech(Numcrop As Integer)
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
    Console.Error.WriteLine((msg))
        msg = "seule possibilité où typinstal=3 et repiquageON=Oui : il y a démariage à irepiqu (réduction de la densité)"
    Console.Error.WriteLine((msg))
    End If
    If TypeInstal(Numcrop) = 1 And ilev(Numcrop) <> iplt(Numcrop) + 1 Then CoherenceIlev = False

    If Not CoherenceIlev Then
        msg = "pb de coherence date de levée, typinstal, date semis"
        msg = msg & "si semis graine prégermée, typinstal=1 et datelevée doit être = isem+1"
    Console.Error.WriteLine((msg))
    End If
    If RepiquageON(Numcrop) = False And densrepiqu(Numcrop) <> DensSem(Numcrop) Then CoherenceDensite = False
    If Not CoherenceDensite Then
        msg = " pb de coherence densité de repiquage / densité de semis "
    Console.Error.WriteLine((msg))
    End If
    End Sub
    Public Sub CyberPlouck(icult As Integer, Die As Boolean, Deathday As Integer, Joursim As Integer, StartDay As Integer)
    'automate ajustant des actes techniques aux conditions du milieu
    If (Die And (Deathday > 0)) And (Joursim + StartDay - 1 < DbutoirNouvSemis(icult)) Then
        Ressemis(icult) = True

    End If
    End Sub
    Public Sub SemisAuto(icult As Integer, Joursim As Integer, StartDay As Integer, Cumpluinf As Double)
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
    Public Sub CompteSemis(icult As Integer)
    Nbsemis(icult) = Nbsemis(icult) + 1
    End Sub
    Public Function sIdCultivar(icult As Integer) As String
    Return IdCultivar(icult)
    End Function
    Public Function nNbCult() As Integer
    Return Nbcult
    End Function
    Public Function niplt(icult As Integer) As Integer
    Return iplt(icult)
    End Function
    Public Function nilev(icult As Integer) As Integer
    Return ilev(icult)
    End Function
    Public Function nirepiqu(icult As Integer) As Integer
    Return irepiqu(icult)
    End Function
    Public Function sTypeInstal(icult As Integer) As Integer
    Return TypeInstal(icult)
    End Function
    Public Function bRepiquageON(icult As Integer) As Boolean
    Return RepiquageON(icult)
    End Function
    Public Function dDensSem(icult As Integer) As Double
    Return DensSem(icult)
    End Function
    Public Function dDensRepiqu(icult As Integer) As Double
    Return densrepiqu(icult)
    End Function
    Public Function bSerreTunnelON() As Boolean
    Return SerreTunnelON
    End Function
    Public Function nCodParamMulch() As Integer
    Return CodParamMulch
    End Function
    Public Function nimulch() As Integer
    Return imulch
    End Function
    Public Function dQpaillisApport() As Double
    Return QpaillisApport
    End Function
    Public Function nAltiCult() As Integer
    Return AltiCult
    End Function
    Public Function bIrrigON() As Boolean
    Return IrrigON
    End Function
    Public Function bfertiminON() As Boolean
    Return fertiminON
    End Function
    Public Function bfertiorgON() As Boolean
    Return fertiorgON
    End Function
    Public Function dtApportMON() As Double
    Return tApportMON
    End Function
    Public Function bRessemis(icult As Integer) As Boolean
    Return Ressemis(icult)
    End Function
    Public Function nJourNouvSem(icult As Integer) As Integer
    Return JourNouvSem(icult)
    End Function
    Public Function nNbsemis(icult As Integer) As Integer
    Return Nbsemis(icult)
    End Function
    Public Function dtApportMinN() As Double
    Return tApportMinN
    End Function
    Public Function bDriveRuiObs() As Boolean
    Return DriveRuiObs
    End Function
    Public Function dCsurN_AO() As Double
    Return CsurN_AO
    End Function
    Public Function dKresY() As Double
    Return KresY
    End Function
End Class
