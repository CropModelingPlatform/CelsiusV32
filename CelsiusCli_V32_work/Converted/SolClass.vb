Option Strict Off
Option Explicit Off
Imports System
Public Class SolClass
    'VX: Variable explicative, VE: variable d'état, VS: variable simulée,PX: paramètre
    Dim IdSol As String
    Dim typsol As String
    Dim NbCouches As Integer
    Dim epc(5) As Integer 'VX, epaisseur de la couche de sol (cm) (sur 5 couches popssibles au maximum)
    Dim hcc(5) As Double 'VX, humidité pondérale à la capacité au champ pour la couche (%)
    Dim hmin(5) As Double 'VX, humidité pondérale au point de flétrissement permanent pour la couche (%)
    Dim da(5) As Double 'VX, densité apparente (sd)
    Dim Ztotsol As Integer 'VX, profondeur totale de sol considérée (somme des epc)
    Dim Eos As Double 'VE, evaporation potentielle du sol (sous mulch et cultures)
    Dim Esol(731) As Double ' VE, evaporation sol du jour(mm)
    Dim Zsurf As Integer 'VX, épaisseur de l'horizon concerné par l'évaporation (cm)
    Dim Stsurf(731) As Double 'VE, stock hydrique disponible de la couche de sol concernée par l'évaporation (mm)
    Dim Strac(731) As Double 'VE, stock hydrique disponible dans la couche de sol colonisée par les racines (mm)
    Dim Stnonrac(731) As Double 'VE, stock hydrique disponible dans la couche de sol non encore colonisée par les racines (entre Plante.Zrac et Plante.Zracmax)(mm)
    Dim Stprofond(731) As Double 'VE, stock hydrique disponible dans la couche de sol non colonisable par les racines (entre Plante.Zracmax et Ztotsol)(mm)
    Dim Stocksol(731) As Double 'VE, Stock hydrique total disponible de 0 à Ztotsol (mm), = Strac+Stnonrac+Stprofond
    Dim Drprofmax(731) As Double 'VS, Drainage sous la cote Ztotsol (mm)
    Dim Dr(731) As Double 'VE, drainage sous la cote maxi colonisable par les racines (Plante.Zracmax) (mm)
    Dim TAW As Double 'VX, ** attention: version par cm de sol de "Total available water in the root zone, mm, as in FAO bull #56 p162" donc ici mm/cm
    ' TAW est la moyenne sur le profil donc si plusieurs couches dans soil_layers elles seront traitées comme une seule de caractéristique
    ' moyenne des couches. TAW est la capacité de stockage moyenne du sol par cm
    Dim TEW As Double 'VX, total Evaporable water, as in FAO bull #56 p144 (mm)
    Dim ZoneHumSousRacines As Double 'VE, comme son nom l'indique, mm
    Dim ContrainteW As Double 'VE, Strac(joursim) / (TAW * Zrac), sd (mm dispo /mm de capacité)
    Dim ContrainteWSurf As Double 'VE, Stsurf(joursim) / TEW, sd (mm dispo sur mm de capacité)
    Dim Zmes As Integer 'VS, profondeur de comparaison des stocks simulés avec des mesures (= profondeur maxi de mesure, par exemple), doit etre < Ztotsol et > Zsurf(cm)
    Dim StockMes(731) As Double 'VS, Stock hydrique total disponible de 0 à Zmes (mm)
    Dim Stger(731) As Double 'VE, Stock hydrique total disponible de 0 à Zger (mm)

    Dim SigmaSimEsol As Double 'VS, somme sur la simulation de Esol (mm)
    Dim SigmaCultEsol As Double 'VS, somme entre semis première culture et récolte dernière culture de Esol (mm)

    Dim SigmaSimDrprofmax As Double 'VS, somme sur la simulation de Drprofmax (mm)
    Dim SigmaCultDrprofmax As Double 'VS, somme entre semis première culture et récolte dernière culture de Drprofmax (mm)
    Dim SigmaSimDr As Double 'VS, somme sur la simulation de Dr (mm)
    Dim SigmaCultDr As Double 'VS, somme entre semis première culture et récolte dernière culture de Dr (mm)
    Dim Zger As Integer 'VX, épaisseur de l'horizon concerné par la germination (cm), transmis au sol à partir d'une lecture faite dans plante

    Dim rstSolData As ADODB.Recordset 'les tables de variables sol lues
    Dim SeuilEvap As Double 'VX, seuil de contrainteWsurf limitant évaporation
    Dim ZObstacleRac As Integer 'VX profondeur (cm) d'un éventuel obstacle à la croissance racinaire
    ' intérêt de passer variable suivante en vecteur du temps ???
    Dim ContrainteHlevee As Boolean 'VE contrainte hydrique pour la germination et la levee
    Dim StockN As Double, StockN1 As Double 'Vx stock du sol en azote organique (kg/ha/an) -  début et fin de simulation
    Dim StockC As Double, StockC1 As Double 'Vx stock du sol en carbone (kg/ha/an) -  début et fin de simulation
    Dim StNJ(731) As Double 'Vx stock du sol en azote (kg/ha/an) -  version journalière
    Dim StCJ(731) As Double 'Vx stock du sol en carbone (kg/ha/an) -  version journalière
    Dim CsurNhum As Double 'Vx C/N ration of Soil Organic Matter (read in Soil table)
    Dim StockP As Double 'Vx stock du sol en phosphore (kg/ha/an) - en attente si developpement de type FIELD ou autre
    Dim StockK As Double 'Vx stock du sol en potassium (kg/ha/an)- en attente si developpement de type FIELD ou autre
    'nouvelles variables bilans N et C
    Dim Cmintot As Double, Nmintot As Double 'VE C et N minéralisés totaux au cours de la saison  en kg.ha-1
    Dim CminMOSY As Double, CminRes As Double, NminMOSY As Double, NminRes As Double 'VE C et N minéralisés sur la saison à partir de la MOS (-MOSY) et des amendements organiques (-Res)
    Dim CminMOSJour(731) As Double, CminResJour(0 To 731) As Double ''VE C minéralisés par jour  à partir de la MOS (-MOSJour) et des amendements organiques (-ResJour)
    ' a vérifier intérêt de passer variable suivante en vecteur du temps ?
    Dim SignalFletrissement(731) As Integer  'VE jours consécutifs accumulés pendant lesquels le sol est au point de flétrissement
    Dim TypeRui As Integer
    Dim TypeSurf As New TypeSurfClass

    Dim VolEausol As Double  'volume d'eau dans Volsol en mètres cubes = Volsol * RUsol (avec RUsol en m/m)
    Dim ConcNsol(731) As Double 'concentration en azote de l'eau du sol kg N/10 mètres cubes d'eau
    Dim ConcNsolmoy As Double 'concentration en azote de l'eau du sol kg N/10 mètres cubes d'eau calculé dans le cas du bilan saisonnier
    Dim DrainageON(731) As Boolean  'si vrai il y a eu drainage sous Zracmax le joursim
    ' parametres du bilan N et C saisonnier
    Dim Clay As Double ' Vx taux d'argile du sol en % - A LIRE
    Dim CaCO3 As Double ' Vx taux de calcaire du sol en % - A LIRE


    Dim txminN As Double 'VE n'est plus lu mais calculé, taux de minéralisation moyen du sol par saison en % du stock d'azote présent - pour bilan saisonnier seulement !
    Dim Fmin As Double, FminY1 As Double 'Vx valeurs début et fin de simul de Fraction minéralisable de la matière organique du sol pour bilan N journalier ET saisonnier = 1-Finert = 0.35 par défaut
    Dim pHeau As Double 'Vx pH eau du sol bour bilan N journalier et saisonnier (nitrification, à faire)

    ' parametres du bilan N journalier (module en travaux, pas évalué)
    Dim perteNDrain(731) As Double  'pertes d'azote par drainage kg/ha
    Dim CumPerteNdrain As Double 'pertes d'azote par drainage kg/ha cumulées sur la simulation
    Dim CumPerteNdrainCult As Double 'pertes d'azote par drainage kg/ha cumulées sur la période semis maturité
    Dim Navail(731) As Double 'stock d'azote disponible en kg/ha
    Dim Tsoil As Double ' VE, temperature du sol dans l'horizon de minéralisation de la MOS (pour l'instant identique à température de l'air). Bilan N Journalier seulement
    Dim Nmin(731)  As Double 'VE, quantité de N du sol minéralisée chaque jour (Kg/ha)
    Dim CNmin(731) As Double  ' VE, Cumul de la quantité de N minéralisé jusqu'au jours de la simulation (Kg/ha).
    Dim MNORG As Double ' VE azote minéralisable de la matière organique du sol pour bilan N journalier
    Dim NDenit(731) As Double  'VE, Quantité d'azote perdu par dénitrification chaque jour à partir du stock de N disponible (Navail)(Kg/ha)
    Dim CNDenit(731) As Double  'VE, Cumul de la quantité d'azote perdu par dénitrification jusqu'au jour de la simulation (Kg/ha).
    Dim MinNorgInput(731) As Double  'VE, Minéralisation journalière de l'azote de la MO apporté(Kg/ha)


    Dim Cres(731) As Double 'C des résidus organiques apportés
    Dim Cbio(731) As Double 'C de la biomasse microbienne
    Dim Nbio(731) As Double 'N de la biomasse microbienne
    Dim Nhum(731) As Double 'N de l'humus

    Dim DNMinOrg(731) As Double 'N de la matière organique apportée (résidus, fumiers)
    Dim CumDNMinOrg(731) As Double ' cumul du N de la matière organique apportée (résidus, fumiers)
    Dim Nreliquat As Double 'VS reliquat azote minéral fin de simul (kg/ha)
    Dim BilanNnonOK As Boolean 'controle si bilan saisonnier en N boucle (False) ou si pb avec N lixivié surestimé (true)
    Dim JourMin As Integer 'VS nombre de jour avec ContrainteWsurf >0.2 pendant la simul
    Dim PRedFact As Double 'Vx Phosphorus constraint (potential conversion efficiency reduction factor) 0-1
    '
    '
    '
    'Dim KR(0 To 731) As Double  'VE Taux de décomposition de la MO apportée de culture pour bilan N journalier
    'Dim FTR(0 To 731) As Double  'VE, Facteur de correction de  la température pour le calcul de la minéralisation de l'N de la MO apportée.
    '
    ' A vérifier:
    '
     ' un critère sur Zrac pour le calcul du flétrissement ?
     ' correction initialisation TEW ne fonctionne pas si sol multicouche
     '
     'VERSION 3 du 8-10 nov 2017
     '
     'modif FA 27/10/2017 calcule stock N à, partir Nstock sol de la table soil en % (da doit être correctement renseigné !)
     ' suite modif 27/10/2017: introduction TxminN lu dans table soil (taux minéralisation moyen saisonnier)
     'modifs en cours: 17/11/2017: introduction de la minéralisation de l'N, introduction de Tsoil,
     ' modif FA du 7/09/18: cumul evap sur culture
     '
     'modifs FA 19/11/21... nouveau bilan saisonnier N et C
     ' attention reste des bouts en chantier...voir ********
     'FA jan 22 modif de eausol pour introduction calcul de  drainage sous profmax pendant période de culture
     'FA aout 22 correction de BilanFin, introduction d'un coef de dilution de l'azote dans les eaux de drainage, proportionnel à la profondeur de la zone racinaire max
     ' ajout de PredFact
     ' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026
    Public Sub LisSol(DataBase_Cnn As ADODB.Connection, SimUnit As SimulationUnitClass, Zgraine As Integer, DriveRui As Boolean)
    Dim Trouve As Boolean
    Dim Icouche As Integer
    Dim msg As String
    Dim reste As Integer
    rstSolData = New ADODB.Recordset

    Trouve = False

    rstSolData.Open("Soil", DataBase_Cnn)
    rstSolData.MoveFirst
    While Not rstSolData.EOF And Not Trouve
        If rstSolData("idSoil") = SimUnit.sIdSoil Then

            Trouve = True
            NbCouches = rstSolData("NbCouches")
            Zsurf = rstSolData("Zsurf")
            Zmes = rstSolData("Zmes")
            SeuilEvap = rstSolData("SeuilEvap")
            ZObstacleRac = rstSolData("ZObstacleRac")
             StockN = rstSolData("StockN")
           ' StockP = rstSolData("StockP")
           ' StockK = rstSolData("StockK")
            CsurNhum = rstSolData("CsurNhum")

            TypeRui = rstSolData("TypeRui")
          '  txminN = rstSolData("txminN") n'est plus lu mais calculé
            Fmin = rstSolData("Fmin")
            Clay = rstSolData("Clay")
            pHeau = rstSolData("pHeau")
            PRedFact = rstSolData("PRedFact")
        End If
        rstSolData.MoveNext
    End While
    rstSolData.Close
    ' message mise en garde ci-dessous à réactiver ?
    'If TypeRui <> 1 And DriveRui Then
     '   msg = "Attention, Typerui conduit à calculer un ruissellement et vous avez choisi un forçage par Ruisselements observés."
     '   msg = msg & " le ruissellement retenu sera les observés pour les jours où il y en a (si pluie >0 ) et les calculés pour les autres"
     '   MsgBox (msg)
    'End If
    Zger = Zgraine
    'affectation de Zgraine lu dans Plante à Zger

    ' lecture données prévue par horizons...mais jamais testée autrement qu'avec un seul horizon
    rstSolData = New ADODB.Recordset
    rstSolData.Open("SELECT * FROM Soil_layers where IdSoil='" & SimUnit.sIdSoil & "' Order by NumCouche", DataBase_Cnn)
    rstSolData.MoveFirst
     For Icouche = 1 To NbCouches
      If Not rstSolData.EOF Then
        epc(Icouche) = rstSolData("epc")
        hcc(Icouche) = rstSolData("hcc")
        hmin(Icouche) = rstSolData("hmin")
        da(Icouche) = rstSolData("da")
    '
        TAW = TAW + (hcc(Icouche) - hmin(Icouche)) * da(Icouche) * epc(Icouche) / 10
        Ztotsol = Ztotsol + epc(Icouche)


             If Icouche = 1 Then
            TEW = (hcc(1) - hmin(1)) * da(Icouche) * Zsurf / 10
            If StockN = 0 Then
                msg = "attention StockN=0; Simulation executée avec StockN=0.005%"
    Console.Error.WriteLine((msg))
                StockN = 0.005
            End If
            StockN = 10000# * StockN * da(1) * 3  ' modif du 27/10/17 N en kg/ha de N à partir de la teneur en %, en considérant
                                           '30cm de sol ds lequel le N orga se minéralise (N/100)x0.3x10000xdax1000=3Nda x10e6/100=3Ndax10000
            StockC = StockN * CsurNhum
            Else
       '
       ' attention ci-dessous pas testé si plusieurs couches et ne parait pas fonctionnel, TEW est intégré de zero à Zsurf
                If Zsurf >= epc(1) And Zsurf > Ztotsol Then
                    reste = Zsurf - Ztotsol
                    TEW = TEW + (hcc(Icouche) - hmin(Icouche)) * da(Icouche) * Min(epc(Icouche), reste) / 10
                End If
            End If
        rstSolData.MoveNext
       Else
       msg = "pb de cohérence nbre de couches tables Soil_layers et Soil"
    Console.Error.WriteLine((msg))
      End If
     Next Icouche
    'TAW en mm/cm (TEW en mm pour la zone 0-Zsurf)
    TAW = TAW / Ztotsol
    'ligne qui suit sera à déplacer si reprise module journalier
    Nhum(0) = Fmin * StockN ' initialisation de l'azote minéral du sol à zero pour bilan journalier de N, impose de démarrer avec un sol sec...?
    '********


    If TEW = 0 Then
    msg = "pb de calcul de TEW; TEW =0"
    Console.Error.WriteLine((msg))
    End If
    If TAW = 0 Then
    msg = "pb de calcul de TAW; TAW =0"
    Console.Error.WriteLine((msg))
    End If
    If Zmes > Ztotsol Or Zmes < Zsurf Then
    msg = "Attention Zmes > Ztotsol Ou Zmes < Zsurf"
    Console.Error.WriteLine((msg))
    End If


    '
    'fermeture table et libération mémoire de l'objet

    rstSolData.Close
    rstSolData = Nothing
    Call TypeSurf.readSurf(DataBase_Cnn, TypeRui)

    End Sub
    Public Sub SeasonMinNC(Nfresh As Double, CsNfresh As Double, KresY As Double, SMT As Double, Napportmin As Double)
    Dim K2humS As Double, fpH As Double, Cfresh As Double, CtoMO As Double, NtoMO As Double
    Dim NhumY0 As Double, ChumY0 As Double, NhumY1 As Double, ChumY1 As Double
    Const K2humref As Double = 0.05

    ' Henin Dupuis en assumant CaCo3 =0% , modifié avec compartiment stable de la MOS, et avec ajout d'une fonction pH ralentissant la nitrification
    ' lorsque le pH est inferieur à 7
    ' KresY = taux de minéralisation saisonnier de la matière organique fraiche apportée
    'Stock C et N initialisés à partir de valeur de StockN et CsurN lue dans sol ou repris séparément des valeurs de la simulation antérieure si simulation récursive (sans passer par CsurN)
    K2humS = K2humref * (1 + 0.2 * (SMT - 10)) / (1 + (0.005 * Clay))
    If pHeau > 7 Then fpH = 1 Else fpH = 0.25 * (pHeau - 3)
    K2humS = K2humS * fpH

    Cfresh = Nfresh * CsNfresh
    ' minéralisation de la matière oraganique fraiche apportée l'année en cours
    ' N et C minéralisés
    NminRes = KresY * Nfresh
    CminRes = KresY * Cfresh
    'C apporté au stock de MO du sol
    'alternative à regarder pour forcer à CsurNmos=10: CtoMO=NtoMO*10 et CminRes= Cfresh - CtoMO ...?

    CtoMO = Cfresh * (1 - KresY)
    NtoMO = Nfresh * (1 - KresY)
    'effect du pH à introduire ? sur N seulement et pas C ? que devient l'ammonium accumulé ?
    'K2hum = K2hum * fpH
    'Mineralisation de la MO du sol
    NhumY0 = Fmin * StockN
    ChumY0 = Fmin * StockC
    NminMOSY = NhumY0 * K2humS
    CminMOSY = ChumY0 * K2humS
    'bilans fin de saison

    Cmintot = CminMOSY + CminRes
    Nmintot = NminMOSY + NminRes
    NhumY1 = NhumY0 - NminMOSY + NtoMO
    ChumY1 = ChumY0 - CminMOSY + CtoMO
    StockN1 = NhumY1 + (1 - Fmin) * StockN
    StockC1 = ChumY1 + (1 - Fmin) * StockC
    'FminY1 = NhumY1 / StockN1
    '
    ' fonction à améliorer pour stabiliser la MO du sol quand apport de MO fraiche à l'humus sont importants ?
    FminY1 = ChumY1 / StockC1
    ' taux moyen de minéralisation calculéà la fin
    txminN = 100 * NminMOSY / StockN
    'CsurNMOS1 = StockC1 / StockN1

    End Sub
    Public Sub initsol(Zracmax As Integer, StockIni As Double, IniSolhautON As Boolean)
    Dim msg As String
    ' 19/12/21 correction bug mineur StockIni >= TAW * Ztotsol  remplacé par StockIni > TAW * Ztotsol

        ContrainteHlevee = True
        StockMes(0) = Min(StockIni, TAW * Zmes)
        If StockIni > TAW * Ztotsol Then
            msg = "attention vous avez mis un stock initial >= capacité de stockage du sol, l'excès ne sera pas compté en drainage"
    Console.Error.WriteLine((msg))
            Stprofond(0) = TAW * (Ztotsol - Zracmax)
            Stnonrac(0) = TAW * Zracmax
            Stsurf(0) = TEW
            Stger(0) = TAW * Zger
        Else
            If IniSolhautON Then

                Stsurf(0) = Min(StockIni, TEW)
                Stger(0) = Min(StockIni, TAW * Zger)
                If StockIni > TAW * Zracmax Then
                    Stnonrac(0) = TAW * Zracmax
                    Stprofond(0) = StockIni - TAW * Zracmax
                Else
                    Stnonrac(0) = StockIni
                End If

            Else
                If StockIni > TAW * (Ztotsol - Zracmax) Then
                    Stprofond(0) = TAW * (Ztotsol - Zracmax)
                    Stnonrac(0) = StockIni - Stprofond(0)
                    If Stnonrac(0) > TAW * Zracmax - TEW Then
                        Stsurf(0) = Stnonrac(0) - (TAW * Zracmax - TEW)
                    End If
                    If Stnonrac(0) > TAW * (Zracmax - Zger) Then
                        Stger(0) = Stnonrac(0) - (TAW * (Zracmax - Zger))
                    End If
                Else
                    Stprofond(0) = StockIni
                End If
            End If
        End If


    Stocksol(0) = Stprofond(0) + Stnonrac(0)

    Navail(0) = 0



    End Sub
    Public Sub InitSolRecurs(EtatFinal As EtatFinalClass)
    Stsurf(0) = EtatFinal.dStsurf_fin
    Stger(0) = EtatFinal.dStger_fin
    Stnonrac(0) = EtatFinal.dStnonrac_fin
    Stprofond(0) = EtatFinal.dStprofond_fin
    Stocksol(0) = EtatFinal.dStocksol_fin
    StockMes(0) = EtatFinal.dStockMes_fin
    Strac(0) = EtatFinal.dStrac_fin
    Fmin = EtatFinal.dFminY1_fin
    StockC = EtatFinal.dStockC1_fin
    StockN = EtatFinal.dStockN1_fin
    End Sub
    Public Sub evaporation(Joursim As Integer, EoSM As Double, Eomulch As Double, pousse As Boolean)
    ' ancienne version : Sub evaporation(joursim As Integer, EoSM As Double, Eomulch As Double, ContrainteWSurf As Double, pousse As Boolean)
    '


    Eos = EoSM - Eomulch

    If ContrainteWSurf < SeuilEvap Then
        Esol(Joursim) = Eos * ContrainteWSurf / SeuilEvap
    Else
        Esol(Joursim) = Eos
    End If
    ' Attention correction de Esol en fonction du niveau d'eau du jour précédent ici dans version 3 et non plus dans Eausol apres calcul de StSurf
    Esol(Joursim) = Min(Stsurf(Joursim - 1), Esol(Joursim))
    SigmaSimEsol = SigmaSimEsol + Esol(Joursim)
    If pousse Then
    SigmaCultEsol = SigmaCultEsol + Esol(Joursim)
    End If
    ' Calcul de SigmasimEsol est à faire dans Eausol si on y déplace la correction de Esol en fonction du stock de surface


    End Sub
    Public Sub TempSoil(T As Double)
    ' en attente d'un effet du mulch ou du LAI -
    Tsoil = T

    End Sub
    Public Sub MinNorgSS(Joursim As Integer)
    'modele de mineralisation du N emprunté à Soltani et Sinclair, modelling physiology of crop development, growth and yield, CABI, 2012
    ' pas testé
    Dim Tfact As Double
    Dim KnMin As Double
    Dim Wfact As Double

    If Tsoil > 35 Then Tsoil = 35
    KnMin = 24 * Exp(17.753 - 6350.5 / (Tsoil + 273))
    Tfact = 1 - Exp(-KnMin / 168)

    If ContrainteWSurf < 0.9 Then
        Wfact = 1.111 * ContrainteWSurf
    Else
        Wfact = 10 - 10 * ContrainteWSurf
    End If
    If Wfact < 0 Then Wfact = 0
    Nmin(Joursim) = Nhum(Joursim - 1) * Tfact * Wfact
    Nmin(Joursim) = ((Nmin(Joursim) * (0.0002 - ConcNsol(Joursim)) / 0.0002)) * 10

    If Nmin(Joursim) < 0 Then Nmin(Joursim) = 0
    Nhum(Joursim) = Nhum(Joursim - 1) - Nmin(Joursim)

    If Joursim = 0 Then
       CNmin(Joursim) = Nmin(Joursim)
    Else
       CNmin(Joursim) = (CNmin(Joursim - 1) + Nmin(Joursim))
    End If
    'If CNmin(Joursim) Then MsgBox ("MinNorg:Ok")
    End Sub
    Public Sub DenitNavail(Joursim As Integer)
    'modele de dénitrification du N du sol emprunté à Soltani et Sinclair, modelling phsiollogy of crop development, growth and yield, CABI, 2012
    'pas testé
    Dim XConcNsol As Double
    Dim KDenit As Double

    NDenit(Joursim) = 0
    If ContrainteWSurf = 1 Then XConcNsol = ConcNsol(Joursim)
    If XConcNsol > 0.0004 Then XConcNsol = 0.0004
    KDenit = 6 * Exp(0.07735 * Tsoil - 6.593)
    NDenit(Joursim) = XConcNsol * (1 - Exp(-KDenit))
    NDenit(Joursim) = (NDenit(Joursim) * Stsurf(Joursim) * 1000) * 10


    If Joursim = 0 Then
    CNDenit(Joursim) = NDenit(Joursim)
    Else
    CNDenit(Joursim) = CNDenit(Joursim - 1) + NDenit(Joursim)
    End If
    End Sub
    Public Sub MinMOSStics(Joursim As Integer)
    Dim Fmin1 As Double 'constante de minéralisation de la MOS du sol
    Dim Fmin2 As Double 'parametre d'effet de la teneur en argile sur la minéralisation de la MOS
    Dim Fmin3 As Double 'parametre d'effet de la teneur en calcair sur la minéralisation de la MOS
    Dim FTH As Double ' fonction de la température du sol
    Dim FH As Double ' fonction de l'humidité du sol
    Dim K2hum As Double 'taux de minéralisation potentiel
    Dim VminH As Double
    'VminH, Nhum(j) à déclarer en global
    'clay et calc = global ? ou K2hum à calculer lors de l'initialisation du sol ?
    'Seulement tyres partiellement testé

    Fmin1 = 0.0006
    Fmin2 = 0.0272
    Fmin3 = 0.0167

    FH = 0.2 + 0.8 * ContrainteWSurf 'attention version stics3
    FTH = 25 / (1 + 145 * Exp(-0.12 * Tsoil))
    K2hum = Fmin1 * Exp(-Fmin2 * Clay) / (1 + Fmin3 * CaCO3)
    VminH = Nhum(Joursim - 1) * K2hum * FH * FTH
    Nhum(Joursim) = Nhum(Joursim - 1) - VminH
    Nmin(Joursim) = VminH
    If Joursim = 0 Then
       CNmin(Joursim) = Nmin(Joursim)
    Else
       CNmin(Joursim) = (CNmin(Joursim - 1) + Nmin(Joursim))
    End If

    End Sub
    Public Sub MinNorgapporteStics(Joursim As Integer, Qorga As Double, CsurNRes As Double, ApportOrga As ApportsOrgaClass)
    'Modèle de minéralisation du N de la MO apportée emprunté à STICS (livre rouge 2012)
    ' attention pas au point: procedure jamais appellée, erreurs de calcul à vérifier
    ' Lectures des parametres des matières organique pas fini du tout, à revoir...
    Dim Kres As Double
    Dim Hres As Double

    Dim FTR As Double ' fonction de la température du sol
    Dim FH As Double ' fonction de l'humidité du sol
    Dim FN As Double ' Fonction de l'azote au voisinage des résidus (désactivée à ce jour)
    Dim CNBio As Double

    Dim DCres As Double 'décomposition journalière du C des résidus
    Dim DNres As Double ' variation journalière du N des résidus
    Dim DCbio As Double 'décomposition journalière du C de la biomasse microbienne
    Dim DNbio As Double 'décomposition journalière du N de la biomasse microbienne
    Dim DChum As Double 'variation journalière du C de l'humus avant minéralisation de ce dernier
    Dim DNhum As Double 'variation journalière du N de l'humus avant minéralisation de ce dernier

    FN = 1
    Const Fbio = 1
    FTR = 12 / (1 + 51.6 * Exp(-0.103 * Tsoil))
    FH = 0.2 + 0.8 * ContrainteWSurf 'attention ici c'est parti de la version Stics3 1998  !
    If Qorga > 0 Then
        If Cres(Joursim - 1) > 0 Then Qorga = ((Cres(Joursim - 1) / CsurNRes) / ApportOrga.dNrec) + Qorga
        Cres(Joursim) = Qorga * ApportOrga.dNrec * CsurNRes
    End If
    Kres = ApportOrga.dAkres + ApportOrga.dBkres / CsurNRes
    Hres = 1 - ApportOrga.dAHres + CsurNRes / (ApportOrga.dBHres + CsurNRes)
    CNBio = ApportOrga.dAWB + ApportOrga.dBWB / CsurNRes
    DCres = -Kres * Cres(Joursim) * FTR * FH * FN
    DNres = DCres / CsurNRes
    DCbio = -ApportOrga.dYres * DCres - ApportOrga.dKbio * Cbio(Joursim - 1) * FTR * FH
    DNbio = -(ApportOrga.dYres * DCres / (CNBio * ApportOrga.dFbio)) - ApportOrga.dKbio * Nbio(Joursim - 1) * FTR * FH
    DChum = ApportOrga.dKbio * Hres * FTR * FH

    DNhum = DChum / CsurNhum
    Cres(Joursim) = Cres(Joursim) + DCbio
    Cbio(Joursim) = Cbio(Joursim - 1) + DCbio
    Nbio(Joursim) = Nbio(Joursim - 1) + DNbio
    DNMinOrg(Joursim) = -DNres - DNhum - DNbio
    If Joursim = 0 Then
       CumDNMinOrg(Joursim) = DNMinOrg(Joursim)
    Else
       CumDNMinOrg(Joursim) = (CumDNMinOrg(Joursim - 1) + DNMinOrg(Joursim))
    End If

    Nhum(Joursim) = Nhum(Joursim) + DNhum ' Nhum(joursim) a déjà été mis à jour 1 premiere fois avec la minéralisation de la MOS
    End Sub
    Public Sub StockNavail(Joursim As Integer, AppMinNjour As Double, NuptMC As Double, Nprecip As Double)
    Const pcentNgaz As Double = 0.01 '% du stock d'azote perdu sous forme de gaz par jour
    'bilan journalier, pas testé, pas appelé
    'Selon Soltani et Sinclair (2012) les Pertes d'N lors de l'application de l'engrais dépendent de la nature de l'engrais apporté, du type d'application (en surface ou incorporé) et de l'environnement (humide, subhumide ou sèche). Par exemple, dans les régions sèches, si l'engrais azoté est appliqué en surface sous forme d'urée juste avant un événement pluvieux, on peut utiliser une pourcentage de perte de 0,25 (25%) .


    'actualisation journalière du stock d'azote disponible
    ' attention pas au point: la minéralisation des apports de MO n'est pas faite (MinNorgINput toujours à 0) et ne doit pas etre appliquée seulement le jour des apports
    ' mais s'appliquer un stock de MO qui est augmenté par les apports et se décompose quotidiennement

    Navail(Joursim) = Navail(Joursim - 1) + Nmin(Joursim) + (AppMinNjour) - (AppMinNjour * 0.35) + DNMinOrg(Joursim) - perteNDrain(Joursim) - NuptMC - (Navail(Joursim - 1) * pcentNgaz) - NDenit(Joursim)

    If Navail(Joursim) < 0 Then Navail(Joursim) = 0
    If Navail(Joursim) > 0 Then
     Navail(Joursim) = Navail(Joursim)
    End If
    End Sub
    Public Sub ConcNEauSol(Joursim As Object, ZracmaxMC As Double)
    'calcul de la concentration en azote de l'eau du sol
    'Le Seuil de ConcNEauSol = 0.0002 gN/g ou 200mgN/l car si ça dépasse cette valeur la minéralisation devient nulle (Nmin (joursim) =0)
    'utilisé pour le calcul des pertes d'azote par les eaux de drainage sous ZracmaxMC

    'calcul du volume d'eau du sol quand il est rempli jusqu'à ZracmaxMC (drainage au-delà de ZracmaxMC) en l/ha
    VolEausol = ZracmaxMC * TAW

    'concentration en azote de l'eau du sol -  Attention on dilue dans l'ensemble du réservoir. Cela devrait etre l'azote est minéralisé dans un réservoir de surface
    ' puis il draine avec la concentration de ce reservoir dans le réservoir sous jacent où il se dilue et est puisé en partie avant drainage éventuel

    ConcNsol(Joursim) = Navail(Joursim) / VolEausol


    End Sub
    Public Sub EauSol(Joursim As Integer, precip As Double, transpi As Double, Zrac As Double, Deltazrac As Double, Zracmax As Integer, pousse As Boolean)

    ' attention dans la forme actuelle la description du sol en couches n'est pas prise en compte
    ' dans le calcul actuel: TAW est considérée comme uniformément répartie
    ' dans le sol en fonction de la profondeur

    Dim bil As Double
    Dim E_Srac As Double
    Dim finrac As Boolean



    bil = Stsurf(Joursim - 1) + precip
    Stsurf(Joursim) = Min(bil, TEW)

    'dans V2 on empêchait ici l'évapopration d'excéder le stock dispo (utile ds cas où seuilevap bas..), mais replacé dans sub evaporation
    ' en fonction stock du jour précédent pour cohérence avec forçage Esol par contrainte du jour précédent et cohérence du calcul du cumul d'éavaporation
    ' Esol(joursim) = Min(Stsurf(joursim), Esol(joursim))
    Stsurf(Joursim) = Stsurf(Joursim) - Esol(Joursim)
    If Zrac > 0 Then Stsurf(Joursim) = Stsurf(Joursim) - Min(Zsurf / Zrac, 1) * transpi
    Stsurf(Joursim) = Max(0, Stsurf(Joursim))
    'attention dans les cas où transpi +evap excède le stock, celui ci est capé à zero mais  transpi du jour
    'n'est pas corrigée donc bilan erroné dans ce cas, en toute rigueur

    ContrainteWSurf = Stsurf(Joursim) / TEW
    ' introduction d'un compteur de jours où sol humide en surface pour calcul du taux de minéralisation journalier moyen à parir du bilan saisonnier
    If ContrainteWSurf > 0.2 Then
        JourMin = JourMin + 1
        Nmin(Joursim) = -1
    End If

    ' introduction du calcul de Stger attention comme on enleve l'evaporation en même temps qu'on remplit, Stger est toujoutrs tres inférieur à Zger*TAW
    If Zrac = 0 Then
        If Zger < Zsurf Then
        Stger(Joursim) = Min(Stger(Joursim - 1) + precip, TAW * Zger) - Esol(Joursim) * Zger / Zsurf
        Else
        Stger(Joursim) = Min(Stger(Joursim - 1) + precip, TAW * Zger) - Esol(Joursim)
        End If
        Stger(Joursim) = Max(0, Stger(Joursim))
    End If
    'contrainte hydrique pour la levée avec des parametres en dur, calés pour mil et arachide au Sénégal !
    If Stger(Joursim) >= 0.14 * TAW * Zger Then ContrainteHlevee = False
    If Stger(Joursim) <= 0.1 * TAW * Zger Then ContrainteHlevee = True


    ' introduction du calcul de StockMes
    If Zrac < Zmes Then
        StockMes(Joursim) = StockMes(Joursim - 1) + precip - Esol(Joursim) - transpi
    Else
        StockMes(Joursim) = StockMes(Joursim - 1) + precip - Esol(Joursim) - transpi * Zmes / Zrac
    'option plus raisonnable pour des simulations de Zmes < Zrac  ?
    'StockMes(joursim) = -999 ' en effet dans ce sol non discrétisé en petites couches
    ' il ne parait pas possible de répartir la transpiration entre deux "tranches" de sol où les racines sont présentes
    End If

    If StockMes(Joursim) < 0 Then StockMes(Joursim) = 0
    If StockMes(Joursim) > TAW * Zmes Then StockMes(Joursim) = TAW * Zmes

    'actualisation du stock dans la zone colonisée par les racines (Stockrac)
    ' traitement du jour de fin de présence d'une plante transpirant par finrac et reportfinrac
    If Zrac = 0 And Deltazrac < 0 Then
        Deltazrac = 0
        finrac = True
    Else
        finrac = False
    End If
    bil = precip + Strac(Joursim - 1) + (Deltazrac * TAW)
    ' Todo: attention on exagère l'acces à l'eau par les racines en croissance...(Deltazrac*TAW)
    ' sera résolu en discrétisant le sol en couches élémentaires...ou en calculant un fronthum

    If bil > Zrac * TAW Then
        Strac(Joursim) = Zrac * TAW
        precip = bil - (Zrac * TAW)
    Else
        Strac(Joursim) = bil
        If Not finrac Then precip = 0 Else precip = precip + Strac(Joursim - 1)
        'si finrac, precip reste inchangé (sera affecté à Stnonrac)
    End If
    'drainage sous racines vers Stnonrac = bil-(Zrac*Taw) si positif
    'si transpi + Evap excède contenu de Strac Strac devient nul
    'TODO: mais il faudrait réduire transpi+évap en conséquence
    ' normalement il n'y a pas de raison de faire le test suivant car evap ne peut excéder Stsurf et tranpi
    ' ne peut excéder Strac

    'Strac(joursim) = Max(0, Strac(joursim))


    'chgt variable pour cas où racines < zsurf
    E_Srac = Esol(Joursim) * Min(1, Zrac / Zsurf)
    Strac(Joursim) = Strac(Joursim) - E_Srac - transpi

    If Strac(Joursim) < 0 Then Strac(Joursim) = 0
    'et dans ce cas y'a une ptite quantité d'eau qui échappe au bilan...
    '

    bil = Stnonrac(Joursim - 1) + precip - (Deltazrac * TAW)

    'si Zrac < Zsurf on affecte la part de l'évaporation correspondante à Snonrac via bil
    If Zrac < Zsurf Then bil = bil - Esol(Joursim) + E_Srac

    If bil > TAW * (Zracmax - Zrac) Then
    Dr(Joursim) = bil - TAW * (Zracmax - Zrac)
    Stnonrac(Joursim) = TAW * (Zracmax - Zrac)
    Stprofond(Joursim) = Min(Stprofond(Joursim - 1) + Dr(Joursim), TAW * (Ztotsol - Zracmax))
    Drprofmax(Joursim) = Max(Stprofond(Joursim - 1) + Dr(Joursim) - TAW * (Ztotsol - Zracmax), 0)
    Else
    Dr(Joursim) = 0
    Stnonrac(Joursim) = Max(0, bil)
    ZoneHumSousRacines = Stnonrac(Joursim) / TAW
    Stprofond(Joursim) = Stprofond(Joursim - 1)
    End If
    Stocksol(Joursim) = Stprofond(Joursim) + Stnonrac(Joursim) + Strac(Joursim)
    If Zrac > 0 Then ContrainteW = Strac(Joursim) / (TAW * Zrac) Else ContrainteW = 1
    ' calcul signal de flétrissement à partir du moment où il y a des racines, avec des "seuils minuscules non nuls" dont il faudrait tester l'influence...
    If Zrac > 0 And Stnonrac(Joursim) <= 0.0001 And Strac(Joursim) <= 0.02 Then

        If SignalFletrissement(Joursim - 1) = 0 Or (Stnonrac(Joursim - 1) <= 0.0001 And Strac(Joursim - 1) <= 0.02) Then
            SignalFletrissement(Joursim) = SignalFletrissement(Joursim - 1) + 1
        End If
    Else
        SignalFletrissement(Joursim) = 0
        ' le signal est remis à zero des que le stock remonte
    End If
    SigmaSimDrprofmax = SigmaSimDrprofmax + Drprofmax(Joursim)
    SigmaSimDr = SigmaSimDr + Dr(Joursim)
    'TODO introduire plage des joursims correspondant à présence culture...
    If pousse Then
        SigmaCultDrprofmax = SigmaCultDrprofmax + Drprofmax(Joursim)
    End If
    'SigmaCultDr = SigmaCultDr + Dr(joursim)

    If Dr(Joursim) > 0 Then DrainageON(Joursim) = True

    'calcul des pertes d'azote par drainage sous Zracmax en kgN/ha
    perteNDrain(Joursim) = ConcNsol(Joursim - 1) * Dr(Joursim)
    CumPerteNdrain = CumPerteNdrain + perteNDrain(Joursim)
    If pousse Then
    CumPerteNdrainCult = CumPerteNdrainCult + perteNDrain(Joursim)
    End If
    End Sub
    Public Sub BilanNFin(Nplante As Double, Napportmin As Double, nbjsimul As Integer, Zracmax As Double)
    Dim J As Integer
    Const FractionDil As Double = 0.01

    StNJ(0) = StockN
    StCJ(0) = StockC
    For J = 1 To nbjsimul
        StNJ(J) = StNJ(J - 1)
        StCJ(J) = StCJ(J - 1)
    'when sub is called, Nmin(J) contains -1 if water is available for mineralisation to occur, 0 otherwise (set in sub EauSol)
        If Nmin(J) = -1 Then
            Nmin(J) = NminMOSY / JourMin
            MinNorgInput(J) = NminRes / JourMin
            CminResJour(J) = CminRes / JourMin
            CminMOSJour(J) = (Cmintot - CminRes) / JourMin
            StNJ(J) = StNJ(J) + (StockN1 - StockN) / JourMin
            StCJ(J) = StCJ(J) + (StockC1 - StockC) / JourMin
        End If
    Next J
    ConcNsolmoy = (Napportmin + Nmintot - Nplante) / VolEausol
    ' la concentration de N à considérer pour l'eau de drainage est fonction décroissante de la profondeur de drainage
    CumPerteNdrain = SigmaSimDr * ConcNsolmoy * Exp(FractionDil * (30 - Zracmax))

    If CumPerteNdrain > Napportmin + Nmintot - Nplante Then
        BilanNnonOK = True
        CumPerteNdrain = Napportmin + Nmintot - Nplante
        Nreliquat = 0
    Else
        Nreliquat = Napportmin + Nmintot - CumPerteNdrain - Nplante
        BilanNnonOK = False
    End If
    If SigmaSimDr > 0 Then CumPerteNdrainCult = (CumPerteNdrain / SigmaSimDr) * SigmaCultDrprofmax Else CumPerteNdrainCult = 0
    End Sub
    Public Function dEsol(Joursim As Integer) As Object
    Return Esol(Joursim)
    End Function
    Public Function dZoneHumSousRacines() As Double
    Return ZoneHumSousRacines
    End Function
    Public Function dContrainteW() As Double
    Return ContrainteW
    End Function
    Public Function dContrainteWSurf() As Double
    Return ContrainteWSurf
    End Function
    Public Function dStsurf(Joursim As Integer) As Double
    Return Stsurf(Joursim)
    End Function
    Public Function dStrac(Joursim As Integer) As Double
    Return Strac(Joursim)
    End Function
    Public Function dStnonrac(Joursim As Integer) As Double
    Return Stnonrac(Joursim)
    End Function
    Public Function dDrprofmax(Joursim As Integer) As Double
    Return Drprofmax(Joursim)
    End Function
    Public Function dStockSol(Joursim As Integer) As Double
    Return Stocksol(Joursim)
    End Function
    Public Function dStprofond(Joursim As Integer) As Double
    Return Stprofond(Joursim)
    End Function
    Public Function iZtotsol() As Integer
    Return Ztotsol
    End Function
    Public Function dStockMes(Joursim As Integer) As Double
    Return StockMes(Joursim)
    End Function
    Public Function dSigmaSimEsol() As Double
    Return SigmaSimEsol
    End Function
    Public Function dSigmaCultEsol() As Double
    Return SigmaCultEsol
    End Function
    Public Function dSigmaSimDrprofmax() As Double
    Return SigmaSimDrprofmax
    End Function
    Public Function dSigmaSimDr() As Double
    Return SigmaSimDr
    End Function
    Public Function nZObstacleRac() As Integer
    Return ZObstacleRac
    End Function
    Public Function bContrainteHlevee() As Boolean
    Return ContrainteHlevee
    End Function
    Public Function dStger(Joursim As Integer) As Double
    Return Stger(Joursim)
    End Function
    Public Function dStockN() As Double
    Return StockN
    End Function
    Public Function dStockP() As Double
    Return StockP
    End Function
    Public Function dStockK() As Double
    Return StockK
    End Function
    Public Function nSignalFletrissement(Joursim As Integer) As Integer
    Return SignalFletrissement(Joursim)
    End Function
    Public Function ClTypeSurf() As Object
    ClTypeSurf = TypeSurf
    End Function
    Public Function dConcNsol(Joursim As Integer) As Double
    Return ConcNsol(Joursim)
    End Function
    Public Function dVolEausol() As Double
    Return VolEausol
    End Function
    Public Function bDrainageON(Joursim As Integer) As Boolean
    Return DrainageON(Joursim)
    End Function
    Public Function dperteNDrain(Joursim As Integer) As Double
    Return perteNDrain(Joursim)
    End Function
    Public Function dDr(Joursim As Integer) As Double
    Return Dr(Joursim)
    End Function
    Public Function dNavail(Joursim As Integer) As Double
    Return Navail(Joursim)
    End Function
    Public Function dCNmin(Joursim As Integer) As Double
    Return CNmin(Joursim)
    End Function
    Public Function dNmin(Joursim As Integer) As Double
    Return Nmin(Joursim)
    End Function
    Public Function dNDenit(Joursim As Integer) As Double
    Return NDenit(Joursim)
    End Function
    Public Function dCNDenit(Joursim As Integer) As Double
    Return CNDenit(Joursim)
    End Function
    Public Function dMinNorgInput(Joursim As Integer) As Double
    Return MinNorgInput(Joursim)
    End Function
    Public Function dNmintot() As Double
    Return Nmintot
    End Function
    Public Function dtxminN() As Double
    Return txminN
    End Function
    Public Function dFminY1() As Double
    Return FminY1
    End Function
    Public Function dStockN1() As Double
    Return StockN1
    End Function
    Public Function dStockC1() As Double
    Return StockC1
    End Function
    Public Function dCmintot() As Double
    Return Cmintot
    End Function
    Public Function dCminRes() As Double
    Return CminRes
    End Function
    Public Function dNminMOSY() As Double
    Return NminMOSY
    End Function
    Public Function dNminRes() As Double
    Return NminRes
    End Function
    Public Function dCumPerteNdrain() As Double
    Return CumPerteNdrain
    End Function
    Public Function dNreliquat() As Double
    Return Nreliquat
    End Function
    Public Function dStockC() As Double
    Return StockC
    End Function
    Public Function dSigmaCultDrprofmax() As Double
    Return SigmaCultDrprofmax
    End Function
    Public Function dCumPerteNdrainCult() As Double
    Return CumPerteNdrainCult
    End Function
    Public Function bBilanNnonOK() As Boolean
    Return BilanNnonOK
    End Function
    Public Function dStNJ(Joursim As Integer) As Double
    Return StNJ(Joursim)
    End Function
    Public Function dStCJ(Joursim As Integer) As Double
    Return StCJ(Joursim)
    End Function
    Public Function dCminMOSJour(Joursim As Integer) As Double
    Return CminMOSJour(Joursim)
    End Function
    Public Function dCminResJour(Joursim As Integer) As Double
    Return CminResJour(Joursim)
    End Function
    Public Function dPRedFact() As Double
    Return PRedFact
    End Function
End Class
