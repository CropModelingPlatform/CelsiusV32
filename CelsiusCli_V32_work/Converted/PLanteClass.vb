Option Strict Off
Option Explicit Off
Imports System
Public Class PLanteClass
    'VX: Variable explicative, VE: variable d'état, VS: variable simulée,PX: paramètre
    Dim rstDataPlante As ADODB.Recordset 'tables "cultivars", "PlantSpecies", "StadePheno"
    Dim Cultivar(2) As String 'identifiant cultivar (lu dans tec_percrop puis recherché dans les tables plantes ici)
    Dim CodeEspece(2) As String 'lu dans la table "cultivars" ici puis recherché dans Plantspecies
    Dim NbStadesPheno(2) As Integer 'nombre de stades phénologiques considérés
    Dim CTstade(2, 10) As Double 'VX, Constante thermique du stade (somme de tempérture seuil de changement de stade)
    Dim TDV(2, 10) As Double 'VS, taux de développement atteint. commence à zero pour TDV(-,stade)-TDV(-,stade-1)
    Dim SensPhot(2) As Double 'VX,Sensibilité à la photopériode
    Dim StrsChoc(2) As Double 'VX, Sensibilité au choc de repiquage (retard en °C.j / unité age de plantules)
    Dim MOPP(2) As Double 'VX, seuil de durée du jour à partir duquel la photopériode agit sur le développement **** FONCTIONNEMENT A VERIFIER
    Dim tdmin(2) As Double
    Dim tdmax(2) As Double 'VX, températures min max et opt de développement (°C). Tdopt présentement non utilisé
    Dim tcmin(2) As Double
    Dim tcmax(2) As Double
    Dim tcopt(2) As Double 'VX, températures min max et opt de conversion de la lumière en biomasse (°C)
    Dim Lai(2, 731) As Double 'VE, LAI journalier du cultivar dans l'association
    Dim Biom(2, 731) As Double 'VS (deviendra VE si prise en compte N), Biomasse aérienne totale journalière du cultivar dans l'association (1000kg/ha)
    Dim Grain(2, 731) As Double 'VS, rendement grain du cultivar dans l'association (T/ha)
    Dim IR(2) As Double 'VX, indice de récolte du cultivar dans l'assoc.
    Dim DLAImax(2) As Double 'VX, croit journalier maximal du lai (sd)
    Dim Currstge(2, 731) As Double 'Stade en cours dans la liste des stades considérés dans la table stadpheno pour le cultivar idCultivar
    Dim DVR(2) As Double 'VE, Development Rate (emprunté à Oryza) du jour, est égal au temps thermique du jour HU (heat unit) divisé par la somme de température du stade en cours
                                ' attention, dans la méthode où DVR est calulé CT désigne l'inverse des constantes thermiques classiques
    Dim DVSt(2, 731) As Double 'VE, cumul de DVR au cours du temps. Comparé à TDV des stades pour identifier si un stade est complété
    Dim TS(2) As Double 'VE, temps thermique cumulé (cumul de HU), version non stockée par jour (°C)
    Dim SommeT(2, 731) As Double 'VE, temps thermique cumulé (cumul de HU), version tableau de stockage par jour (°C)
    Dim TSTR As Double 'VE, temperature sum at transplanting (oryza), sert à calculer l'effet du stress de repiquage: tant que TS < TSTR + stsrschc on ne reprend pas le développement
    Dim DVS(2) As Double 'VE, cumul de DVR version valeur en cours (°C)
    Dim JulPheno(2, 6) As Integer 'VS, date où le stade "n" a été atteint (en jour depuis le1/01 de l'année de début de la simulation)
    Dim cropsta(2, 731) As Integer 'VE, état du cultivar dans l'assoc
                                        '0=before sowing; 1=day of sowing; 2=in seedbed;
    '                                    3=after emergence and before transplanting; 4=main growth period
    Dim Die(2) As Boolean 'VE, état de vie (faux) ou mort (vrai) du cultivar
    Dim Transplant(2) As Boolean 'VX il y a (vrai) où il n'y a pas repiquage/démariage
    Dim Death_day(2) As Integer 'VS, jour de la mort éventuelle de la culture (en jour depuis le 1er jour de simulation)
    Dim JourLaiMax(2) As Integer 'VS, jour d'atteinte du LAI maximal (en jour depuis le 1er jour de simulation)
    Dim JourSen(2) As Integer 'VS, jour du début de sénescence du cultivar (en jour depuis le 1er jour de simulation)
    Dim JourDrp(2) As Integer 'VS jour de début de remplissage des grains pour le cultivar (en jour depuis le 1er jour de simulation)
    Dim JourMat(2) As Integer 'VS, jour de maturité physiologique du cultivar (en jour depuis le 1er jour de simulation)
    Dim Hu(2) As Double 'VE, Heat Units (Oryza), temps thermique en °C jour, présentement calculé dans HUstics (sans effet de Tdopt)
    Dim densplt(2, 731) As Double 'VX, densité de plantes au semis ( /m2)
    Dim adens(2) As Double 'VX, coefficient de sensibilité aux densités élevées (voir Stics)
    Dim bdens(2) As Double 'VX, seuil de densité à partir duquel la surface foliaire par plante dépend de la densité
    Dim Laicomp(2) As Double 'VX, LAI seuil à partir duquel la densité de peuplement peut avoir un impact sur la surface foliaire par plante (au-delà du seuil de densité bdens)
    Dim LAIrec(2) As Double 'VX, LAI résiduel à la récolte
    Dim CoefExtin(2) As Double 'VX, coefficient d'extinction du rayonnement par le LAI
    Dim Ebmax(2) As Double 'VX, epsilon b max, taux potentiel de conversion du rayonnement en biomasse, gMS.MJ(-1)
    Dim Zrac(2, 731) As Double 'cote atteinte par les racines (cm)
    Dim DeltaRacMax(2) As Double ' VX croissance journalière de la profondeur atteinte par les racines par unité de temps thermique (cm/°.j)
    Dim Zracmax(2) As Integer  'VX  Cote maximale atteignable par les racines (cm) ***attention sera réduite à Sol.Ztotsol si Zracmx > Ztotsol
    ' variables et paramètres de la sensibilité au froid
    Dim Ncold(2) As Integer 'VE, nombre de jours consécutifs où la température est inférieure au seuil de sensibilité au froid Tcold
    Dim Tcold(2) As Double 'VX, Température seuil de sensibilité au froid (°C)
    Dim NDieCold(2) As Integer 'VX, nombre maximal de jours inférieurs à Tcold supportable par la culture sans mortalité
    ' fin variables et paramètres de la sensibilité au froid
    Dim deltaBiom(2, 731) As Double
    Dim Vitmoy(2) As Double 'VE, vitesse moyenne de croissance pendant la phase de détermination du nombre de grains (en g/m2/j)
    Dim Nbjgrain(2) As Integer 'VX, nombre de jours déterminant le nombre de grains avant le stade 4 (début remplissage du grain)
    Dim Ngrains(2) As Long 'VE, nombre de grains par m2
    Dim Cgrain(2) As Integer 'VX, nombre de grains mis en place par gMS/jour de croissance moyenne de biomasse pendant les nbjgrain précédent le début du stade 4
    Dim Cgrainv0(2) As Long 'VX, nombre de grains mis en place si croissance nulle pendant les nbjgrain (grains /m2)
    Dim Ngrmax(2) As Long 'VX nombre maximal de grains par plante
    Dim Vitircarb(2) As Double 'VX, augmentation journalière de l'indice de récolte (g grain. g(-1) Ms j(-1))
    Dim IRmax(2) As Double 'VX, indice de récolte maximal du cultivar
    Dim P1grainMax(2) As Double 'VX poids maximal de 1 grain (g)
    Dim P1grain(2, 731) As Double 'VE poids d'un grain (g)
    Dim Kmax(2) As Double 'Vx coefficient cultural max de l'espece quand LAI >5
    Dim TSlevee(2) As Integer 'VE, temps thermique cumulé pour la levée+germination(cumul de HU)
    Dim CTlevee(2) As Integer 'VX constante thermique de germination-levée
    Dim LevSim(2) As Integer 'VE, date de levée simulée en jour de la simulation
    Dim Tger(2) As Double 'VX, températures base pour levée (°C)
    Dim stressN As Double 'VE coefficient de réduction de l'efficience de conversion en fonction du statut azoté de la plante (0 pas de croissance, 1 pas de stress)
    Dim LegumON(2) As Boolean 'VX, légumineuse oui ou non
    Dim Nsymb(2) As Double  'VX, stock de nutriments (azote) fixé par symbiose mycorhizienne
    Dim IFertMax(2) As Double 'azote (kg/ha): seuil de fertilisation au delà duquel l'indice de satisfaction ds besoins nutritifs est égal à 1
    ' a vérifier nouvelles variables introduites par CP
    Dim NCvEmax(2) As Double 'efficience de conversion maximale de l'azote en biomasse (kg MS/kg N)
    Dim NCvEmin(2) As Double 'efficience de conversion minimale de l'azote en  biomasse (kg MS/kg N)
    Dim alphaN(2) As Double 'coefficient de réglage de l'efficience de conversion de l'azote en biomasse
    Dim PCvEmax(2) As Double 'efficience de conversion maximale du phosphore en biomasse (kg MS/kg P)
    Dim PCvEmin(2) As Double 'efficience de conversion minimale du phosphore en  biomasse (kg MS/kg P)
    Dim alphaP(2) As Double 'coefficient de réglage de l'efficience de conversion du phosphore en biomasse
    Dim KCvEmax(2) As Double 'efficience de conversion maximale du potassium en biomasse (kg MS/kg K)
    Dim KCvEmin(2) As Double 'efficience de conversion minimale du potassium en  biomasse (kg MS/kg K)
    Dim alphaK(2) As Double 'coefficient de réglage de l'efficience de conversion du potassium en biomasse
    Dim NRF(2, 731) As Double 'nitrogen reduction factor = dispo/demande
    Dim NRF_lai(2, 731) As Double 'nitrogen reduction factor = dispo/demande pour effet sur LAI
    Dim NRF_bio(2, 731) As Double 'nitrogen reduction factor = dispo/demande pour effet sur biomasse
    Dim PRF(731) As Double 'phosphorus reduction factor = dispo/demande
    Dim KRF(731) As Double 'potassium reduction factor = dispo/demande
    Dim SensiSen(2) As Double 'VX sensibilité de la sénéscence aux stress post floraison
    Dim NJFletri(2) As Integer  'VX nombre de jours consécutifs de sol au point de flétrissement au-delà duquel il ya mort de la culture (dans plantespecies))
    'A verifier, valeurs introduites dans tables Todo a lire dans table
    Dim Zgraine(2) As Integer 'VX, épaisseur de l'horizon concerné par la germination (cm)

    Dim SeuilTurg(2) As Double   ' PX 1-seuil d'effet de la contrainte hydrique sur la croissance du LAI (règle lien entre contrainteH et turfac, typiquement=0.25)
    Dim SeuilWS(2) As Double  'PX, 1-seuil d'effet de la contrainte hydrique sur la croissance de la biomasse(règle lien entre contrainteH et WSfact, typiquement=0.35)
    'ci-dessous variables du module stressAzote
    Dim NavailCult(2, 731) As Double   'stock d'azote disponible en kg/ha en tenant compte de la fixation symbiotique (Navail + Nsymbjour)
    Dim NUPTtarget(2, 731) As Double 'demande en azote (kg/ha), dépend du croît de biomasse du jour sans stress nutritif
    Dim Nuptake(2, 731) As Double  'quantité d'azote réellement absorbée par plante = NUPTtarget si aucun autre facteur limitant
    Dim SigmaNuptake(2, 731)  'cumul de la consommation d'azote par la plante
    Dim Nsymbjour(2) As Double 'quantité d'azote fixée par jour = Nsymb/durée du cyle
    'A vérifier attention CP a vectorisé toutes ces variables - pas sur que ce soit bonne idée (tps calcul !)
    ' facteurs de stress
    Dim TurfacH(731) As Double 'utilisé dans Calcule_LAI
    Dim Turfac(731) As Double 'utilisé dans Calcule_LAI
    Dim WSfactH(731) As Double 'utilisé dans Biomasse
    Dim WSfact(731) As Double 'utilisé dans Biomasse
    Dim Pfactor(2) As Double 'seuil de réduction de la transpiration lorsque la teneur en  eau du sol diminue (maj juin2025)
    'variables en lien avec la compétition pour la lumière
    Dim Competition(731) As Boolean 'si true il y a compétition pour la lumière entre les deux espèces de l'association
    Dim CompFac(2, 731) As Double 'facteur de compétition pour la lumière
    Dim raint(2, 731) As Double
    ' variables effet changement climatique
    Dim alphaCO2(2) As Double 'VX sensibilité de la conversion en biomasse à la concentration en CO2 de l'atmosphere (C3: 1.2; C4: 1.1)
    Dim CO2fact(2) As Double 'VE facteur de réduction par le CO2 de la conversion du rayonnement en biomasse
    Dim LAIpot(2) As Double 'VS lai potentiel maxi qui serait atteint sans stress à julpheno_3
    ' (prévu pour utilisation dans effet des stress post julpheno3 dans calculeLAIsemiaride mais finalement pas utilisé)
    Dim DurCycMax(2) As Integer 'VX durée max du cycle de culture en jours
    Dim Nplant(2) As Double ' VS Azote exporté dans la plante kg/ha
    Dim concNplante(2) As Double 'Vx concentration de reference de N dans la plante kgN / t biomasse
    Dim RootABGRatio(2) As Double
    Dim RootCN(2) As Double
    Dim NRootABGRatio(2) As Double
    Dim NGrainABGRatio(2) As Double
    Dim PlantPReducFact As Double

    Const Kmo = 0.25 'VX pour stresNold part de l'azote apporté par amendements organique minéralisé (net) par an
    'Modif janvier 2020 a partir version 3 du 8-10 nov 2017'
    'procedure introduite stressAzoteOldCourbe avec les parametres ci-dessous
    'Const sensiStressNLAi = 0.015   ' dernier calage fction 1-exp calage initial 0.03
    'Const sensiStressNBio = 0.0025   'dernier calage fction 1-exp calage initial 0.02
    Const sensiStressNLAi = 0.015  'calage initial sigmo:
    Const sensiStressNBio = 0.005   ''calage initial sigmo
    Const InflexStressN = 60 ' point d'inflexion de la sigmoide du stress N
    '
    '
    ' a vérifier:
    ' pertinence de passer en vecteur jour les variables de stress
    ' effet de comptétition sur rayonnement intercepté
    ' effet stress hydrique sur développement
    ' commentaires insérés
    '
    'VERSION 3.1 jan 2020
    ' Modifs FA nov 2021: stressAzoteOld modifié pour utiliser le N mineralisé calculé dans Sol.
    ' nouveau parametre durcycmax pour le cultivar -
    ' modif Aout 22...PlantPReducFact introduit dans PlanteClass, reçu de soil class dans Iniplante et reduisant ebmax dans deltabiom
    'Version 3.2 Jan 2025
    'dlaimax now read in Cultivars instead of PlantSpecies
    'Pfactor lu ds plantSpecies pour etre transmis à TranspiMC
    ' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026
    Public Sub LisPlante(DataBase_Cnn As ADODB.Connection, GestionTechnique As GestionTechniqueClass)
    Dim Trouve As Boolean
    Dim icult As Integer
    Dim Numstade As Integer

    For icult = 1 To GestionTechnique.nNbCult


      rstDataPlante = New ADODB.Recordset
      Trouve = False
      Cultivar(icult) = GestionTechnique.sIdCultivar(icult)
      rstDataPlante.Open("Cultivars", DataBase_Cnn)
      rstDataPlante.MoveFirst
            While Not rstDataPlante.EOF And Not Trouve
                If rstDataPlante("IdCultivar") = Cultivar(icult) Then

                    Trouve = True
                    CodeEspece(icult) = rstDataPlante("CodePSpecies")
                    NbStadesPheno(icult) = rstDataPlante("NbStadesPheno")
                    SensPhot(icult) = rstDataPlante("SensPhot")
                    If GestionTechnique.bRepiquageON(icult) Then
                        StrsChoc(icult) = rstDataPlante("StrsChoc")
                        Transplant(icult) = True
                    Else
                        StrsChoc(icult) = 0
                        Transplant(icult) = False
                    End If
                    MOPP(icult) = rstDataPlante("MOPP")
                    DLAImax(icult) = rstDataPlante("DLAImax")
                    adens(icult) = rstDataPlante("adens")
                    bdens(icult) = rstDataPlante("bdens")
                    Laicomp(icult) = rstDataPlante("Laicomp")
                    Vitircarb(icult) = rstDataPlante("Vitircarb")
                    IRmax(icult) = rstDataPlante("IRmax")
                    P1grainMax(icult) = rstDataPlante("P1grainMax")
                    Ngrmax(icult) = rstDataPlante("Ngrmax")
                    IFertMax(icult) = rstDataPlante("IFertMax")
                    Cgrain(icult) = rstDataPlante("Cgrain")
                    Cgrainv0(icult) = rstDataPlante("Cgrainv0")
                    DurCycMax(icult) = rstDataPlante("DurCycMax")
                    concNplante(icult) = rstDataPlante("concNplante")
                    SensiSen(icult) = rstDataPlante("SensiSen")

                End If
                rstDataPlante.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet
      rstDataPlante.Close
      rstDataPlante = New ADODB.Recordset
      Trouve = False
      rstDataPlante.Open("PlantSpecies", DataBase_Cnn)
      rstDataPlante.MoveFirst
            While Not rstDataPlante.EOF And Not Trouve
                If rstDataPlante("CodePSpecies") = CodeEspece(icult) Then
                    Trouve = True
                    tdmin(icult) = rstDataPlante("tdmin")
                    tdmax(icult) = rstDataPlante("tdmax")
                    tcmin(icult) = rstDataPlante("tcmin")
                    tcmax(icult) = rstDataPlante("tcmax")
                    tcopt(icult) = rstDataPlante("tcopt")
    '               DLAImax(icult) = rstDataPlante("DLAImax") now read in Cultivars modif FA Jan 2025
                    LAIrec(icult) = rstDataPlante("LAIrecmax")
                    CoefExtin(icult) = rstDataPlante("extin")
                    Ebmax(icult) = rstDataPlante("Ebmax")
                    DeltaRacMax(icult) = rstDataPlante("DeltaRacMax")
                    Zracmax(icult) = rstDataPlante("Zracmax")
                    Tcold(icult) = rstDataPlante("Tcold")
                    NDieCold(icult) = rstDataPlante("NDieCold")
                    Nbjgrain(icult) = rstDataPlante("Nbjgrain")
                    Kmax(icult) = rstDataPlante("Kmax")
                    CTlevee(icult) = rstDataPlante("CTlevee")
                    Tger(icult) = rstDataPlante("Tger")
                    LegumON(icult) = rstDataPlante("LegumON")
    '                SensiSen(icult) = rstDataPlante("SensiSen")
    ' sensisen lu dans cultivar, modif du 28/08/22
                    NJFletri(icult) = rstDataPlante("NJFletri")
                    Zgraine(icult) = rstDataPlante("Zgraine")
                    Nsymb(icult) = rstDataPlante("Nsymb")
                    NCvEmax(icult) = rstDataPlante("NCvEmax")
                    NCvEmin(icult) = rstDataPlante("NCvEmin")
                    alphaN(icult) = rstDataPlante("alphaN")
                    PCvEmax(icult) = rstDataPlante("PCvEmax")
                    PCvEmin(icult) = rstDataPlante("PCvEmin")
                    alphaP(icult) = rstDataPlante("alphaP")
                    Nsymbjour(icult) = rstDataPlante("Nsymbjour")
                    SeuilTurg(icult) = rstDataPlante("SeuilTurg")
                    SeuilWS(icult) = rstDataPlante("SeuilWS")
                    alphaCO2(icult) = rstDataPlante("alphaCO2")
                    RootABGRatio(icult) = rstDataPlante("RootABGRatio")
                    RootCN(icult) = rstDataPlante("RootCN")
                    NRootABGRatio(icult) = rstDataPlante("NRootABGRatio")
                    NGrainABGRatio(icult) = rstDataPlante("NGrainABGRatio")
                    Pfactor(icult) = rstDataPlante("Pfactor")
                End If
                rstDataPlante.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet
      rstDataPlante.Close
      rstDataPlante = New ADODB.Recordset
      rstDataPlante.Open("SELECT * FROM StadePheno where CodCultivar='" & Cultivar(icult) & "' Order by NumStade", DataBase_Cnn)
      rstDataPlante.MoveFirst
            While Not rstDataPlante.EOF
                Numstade = rstDataPlante("Numstade")
                 CTstade(icult, Numstade) = rstDataPlante("CTstade")
                 TDV(icult, Numstade) = rstDataPlante("TDV")

                rstDataPlante.MoveNext
    End While
    'fermeture table et libération mémoire de l'objet
      rstDataPlante.Close
    rstDataPlante = Nothing

    Next icult


    End Sub
    Public Sub Iniplante(icult As Integer, DensSem As Double, isem As Integer, densrepiqu As Double, irepiqu As Integer, DebutDOY As Integer, finDOY As Integer, LimiteSol As Integer, CO2c As Integer, PreducFact As Double)

    Dim Joursim As Integer
    Dim jourfin As Integer
    Dim Dens As Double

    PlantPReducFact = PreducFact
    CO2fact(icult) = FCO2(CO2c, icult)

    Die(icult) = True
    If irepiqu <> 999 Then jourfin = irepiqu - DebutDOY + 1 Else jourfin = finDOY
    If Zracmax(icult) > LimiteSol Then Zracmax(icult) = LimiteSol
    If isem - DebutDOY + 1 < 0 Then
    Console.Error.WriteLine(("semis avant le début de la simulation, stop! corrigez TechPerCrop"))
    End If
    For Joursim = isem - DebutDOY + 1 To jourfin - 1
        densplt(icult, Joursim) = DensSem
    Next Joursim
    For Joursim = jourfin To finDOY
        densplt(icult, Joursim) = densrepiqu
    Next Joursim

    If irepiqu <> 999 Then Dens = densrepiqu Else Dens = DensSem

    If Dens <= bdens(icult) Then LAIrec(icult) = LAIrec(icult) * Dens / bdens(icult)
    'initialisations indispensables pour simulation de ressemis après mort de la culture
    TSlevee(icult) = 0
    TS(icult) = 0
    DVS(icult) = 0
    DVR(icult) = 0

    End Sub
    Public Sub AdapteCT(Nbcult As Integer)
    ' conversion des parametres de développement pour saisir des constantes
    ' thermiques en °C.j et conserver le formalisme de Oryza pour l'age physiologique (development stage)
    Dim icult As Integer
    Dim Istade As Integer
    For icult = 1 To Nbcult
        For Istade = 1 To NbStadesPheno(icult)
        CTstade(icult, Istade) = Inverse(CTstade(icult, Istade)) * (TDV(icult, Istade) - TDV(icult, Istade - 1))
        Next Istade
    Next icult
    End Sub
    Public Function HUstics(Tm As Double, icult As Integer) As Double

    Dim TT As Double

    TT = Tm - tdmin(icult)
    If (Tm < tdmin(icult)) Then TT = 0
    If (Tm > tdmax(icult)) Then TT = tdmax(icult) - tdmin(icult)
    HUstics = TT
    'une correction tempps thermique par stress H possible ici sous la forme HUstics=TT*(1+QstressH)
    End Function
    Public Function HUleve(Tm As Double, icult As Integer) As Double

    Dim TTleve As Double

    TTleve = Tm - Tger(icult)
    If (Tm < Tger(icult)) Then TTleve = 0
    HUleve = TTleve

    End Function
    Public Sub GerminLevee(DOY As Integer, icult As Integer, Joursim As Integer, ilev As Integer, simlevee As Boolean, ContrainteHlevee As Boolean, Tm As Double)
    'procédure appellée pour joursim compris entre joursemis et ilev (inclus), via test sur cropsta

    Dim msg As String
    cropsta(icult, Joursim) = 2
    If Not simlevee And ilev <> 0 And ilev <> 999 Then

        If DOY = ilev Then
            If Transplant(icult) Then cropsta(icult, Joursim) = 3 Else cropsta(icult, Joursim) = 4
        End If

    Else
        If Not ContrainteHlevee Then
        TSlevee(icult) = TSlevee(icult) + HUleve(Tm, icult)
            If TSlevee(icult) >= CTlevee(icult) Then
    ' ligne suiv. et variable LevSim pas forcément utile ? A verifier plus systematiquement avec des cas de levée simulée
    '            LevSim(icult) = Joursim
                If Transplant(icult) Then cropsta(icult, Joursim) = 3 Else cropsta(icult, Joursim) = 4
            End If
        End If
    End If
    End Sub
    Public Sub stressAzoteOld(icult As Integer, Joursim As Integer, Nmintot As Double, ApportMin As Double)
    Dim QNut As Double
    QNut = Nmintot + ApportMin + Nsymb(icult)
    stressN = QNut / IFertMax(icult)
    If stressN > 1 Then stressN = 1
    NRF(icult, Joursim) = stressN
    NRF_lai(icult, Joursim) = NRF(icult, Joursim)
    NRF_bio(icult, Joursim) = NRF(icult, Joursim)
    ' utile conserver deux variables NRF et stressN à cause de la variante ci-dessous
    End Sub
    Public Sub ExportN(Nminsol As Double)
    Dim Np1 As Double, Np2 As Double
    Dim icult As Integer

    If Nminsol > 0 Then
        For icult = 1 To 2
            Nplant(icult) = Biom(icult, JourMat(icult)) * concNplante(icult)
        Next icult

        If Nplant(1) + Nplant(2) > Nminsol Then
            Np1 = Nminsol * Nplant(1) / (Nplant(1) + Nplant(2))
            Np2 = Nminsol * Nplant(2) / (Nplant(1) + Nplant(2))
            Nplant(1) = Np1
            Nplant(2) = Np2
        End If
    Else
        Nplant(1) = 0
        Nplant(2) = 0
    End If
    End Sub
    Public Sub stressAzoteOldCourbe(icult As Integer, Joursim As Integer, Nmintot As Double, ApportMin As Double)
    'modifFA 29/11/2019 introduction de deux fonctions de stress différentes pour LAI et biomasse
    Dim QNut As Double

    QNut = Nmintot + ApportMin + Nsymb(icult)
    stressN = QNut / IFertMax(icult)
    If stressN > 1 Then stressN = 1
    ' option 1-exponentielle
    'NRF_lai(icult, Joursim) = 1 - Exp(-sensiStressNLAi * QNut)
    'NRF_bio(icult, Joursim) = 1 - Exp(-sensiStressNBio * QNut)
    ' option sigmoide
    NRF_lai(icult, Joursim) = 1 / (1 + Exp(-sensiStressNLAi * (QNut - InflexStressN)))
    NRF_bio(icult, Joursim) = 1 / (1 + Exp(-sensiStressNBio * (QNut - InflexStressN)))

    End Sub
    Public Sub phenoCTphot(dataclim As DataClimClass, Joursim As Integer, SignalFletrissement As Integer, icult As Integer)
    'Attention NOUVELLE approche plus directe que celle empruntée à Oryza, sans inversion des constantes thermiques.
    ' Verifications en cours. OK sur le jeu 840 Senegal sans photosensibilité, différences mineures (+- 1j sur stade6) avec Sigma_CT si photosensibilite...
    ' verifier fletrissement, froid, repiquage...

    Dim DL As Double 'mort de la culture si le sol est au point de flétrissement depuis plus de NJFletri
    Dim PPFAC As Double
    Dim TxDev As Double

    Dim StopTransplant As Integer

    Currstge(icult, Joursim) = Currstge(icult, Joursim - 1)
    PPFAC = 1

    If SignalFletrissement >= NJFletri(icult) And Currstge(icult, Joursim) < 2 Then
        Die(icult) = True
        Death_day(icult) = dataclim.nDOY(Joursim)
        Currstge(icult, Joursim) = 6 'verifier
        cropsta(icult, Joursim) = 0
    End If

    ' début prise en compte froid
    If dataclim.dTmoy(Joursim) < Tcold(icult) Then Ncold(icult) = Ncold(icult) + 1 Else Ncold(icult) = 0

    If Ncold(icult) = NDieCold(icult) Then
        Die(icult) = True
        Death_day(icult) = dataclim.nDOY(Joursim)
        Currstge(icult, Joursim) = 6 'verifier
        cropsta(icult, Joursim) = 0 'verifier
    End If
    'fin froid

    ' les phases 1 à 5 ci dessous (currstge 1 = croissance lente du LAI, julpheno1= levée; currstge5 = sénéscence feuilles et poursuite du remplissage du grain, julpheno6= maturité, julpheno)
    Hu(icult) = HUstics(dataclim.dTmoy(Joursim), icult)

    ' prise en charge photopériode
    If Currstge(icult, Joursim) = 2 Then
             DL = dataclim.dDAYL(Joursim) + 0.9
             If (DL < MOPP(icult)) Then
                PPFAC = 1#
             Else
                PPFAC = 1# - (DL - MOPP(icult)) * SensPhot(icult)
             End If
             PPFAC = Min(1#, Max(0#, PPFAC))
    End If
    'somme T sans frein photopériodique pour sortie journalière
    SommeT(icult, Joursim) = SommeT(icult, Joursim - 1) + Hu(icult)
    ' Somme T pour calcul du développement avec frein photopériodique
    TxDev = Hu(icult) * PPFAC
    If (cropsta(icult, Joursim - 1) > 3 And TS(icult) + TxDev < (TSTR + StrsChoc(icult))) Then StopTransplant = 0 Else StopTransplant = 1

    'attention StrsChoc doit etre nul si pas de repiquage

    TxDev = TxDev * StopTransplant
    TS(icult) = TS(icult) + TxDev


    'DVS : cumul des constantes thermiques jusqu'au stade en cours inclus (le stade change quand la somme photothermique dépasse DVS)
    If Not Die(icult) And TS(icult) >= DVS(icult) Then
        Currstge(icult, Joursim) = Currstge(icult, Joursim) + 1
        If Currstge(icult, Joursim) = 6 Then
            Die(icult) = True
            JourMat(icult) = Joursim
        Else
    ' augmentation du seuil de temps thermique de la constante thermique du nouveau stade
            DVS(icult) = DVS(icult) + CTstade(icult, Currstge(icult, Joursim))
        End If

    End If
    If Currstge(icult, Joursim) < 6 Then
    'DVSt(icult, joursim): somme des taux de developpement normalisés selon échelle des TDV(stade), pour emploi dans calcul Lai
        DVSt(icult, Joursim) = DVSt(icult, Joursim - 1) + (TDV(icult, Currstge(icult, Joursim)) - TDV(icult, Currstge(icult, Joursim) - 1)) * TxDev / CTstade(icult, Currstge(icult, Joursim))
    End If

    '!-----CROPSTA: 0=before sowing; 1=day of sowing; 2=in seedbed;
    '!                  3=day of transplanting; 4=main growth period

    'attention dans ce qui suit voir s'il faut calculer TSTR par icult si cultures associées...
    If (cropsta(icult, Joursim) = 4) And cropsta(icult, Joursim - 1) = 3 Then TSTR = TS(icult)



    'la date julpheno(icult, j) est la date à laquelle un nouveau currstge(icult) démarre julpheno1 le lendemain du dernier jour de currstge1
    If Currstge(icult, Joursim) > Currstge(icult, Joursim - 1) Then JulPheno(icult, Currstge(icult, Joursim)) = dataclim.nDOY(Joursim)


    End Sub
    Public Sub pheno_sigmaT(dataclim As DataClimClass, Joursim As Integer, SignalFletrissement As Integer, icult As Integer)
    'attention avbec cette routine CT doit être l'inverse des données entrées dans table StadePheno, voir AdapteCT
    Dim DL As Double
    Dim PPFAC As Double
    'mort de la culture si le sol est au point de flétrissement depuis plus de NJFletri

    If SignalFletrissement >= NJFletri(icult) And DVS(icult) < TDV(icult, 2) Then
        Die(icult) = True
        Death_day(icult) = dataclim.nDOY(Joursim)
        DVS(icult) = TDV(icult, 5)
        cropsta(icult, Joursim) = 0
    End If

        If TS(icult) = 0 Then DVR(icult) = CTstade(icult, 1)

    ' a reprendre ce qui suit pour généricité
    Hu(icult) = HUstics(dataclim.dTmoy(Joursim), icult)
    ' redondance TS, SommeT et DVS, DVSt ?
    ' somme de temps thermique (Hu = heat unit= le delta de temps thermique du jour)
    TS(icult) = TS(icult) + Hu(icult)
    'le taux de développement
    DVR(icult) = DVR(icult) * Hu(icult)
    'le stockage du temps thermique dans tableau pour OutputD
    SommeT(icult, Joursim) = TS(icult)
    'le stade de développement en continu
    DVS(icult) = DVS(icult) + DVR(icult)
    ' début prise en compte froid
    If dataclim.dTmoy(Joursim) < Tcold(icult) Then Ncold(icult) = Ncold(icult) + 1 Else Ncold(icult) = 0

    If Ncold(icult) = NDieCold(icult) Then
        Die(icult) = True
        Death_day(icult) = dataclim.nDOY(Joursim)
        DVS(icult) = TDV(icult, 5)
        ' ligne ci-dessus provoque plus loin die=true et DVR=0
    End If
    'fin froid
    ' les phases 1 à 5 ci dessous (currstge 1 = croissance lente du LAI, julpheno1= levée; currstge5 = sénéscence feuilles et poursuite du remplissage du grain, julpheno6= maturité, julpheno)
    DVSt(icult, Joursim) = DVS(icult)
    If (DVS(icult) >= 0 And DVS(icult) < TDV(icult, 1)) Then
        DVR(icult) = CTstade(icult, 1)
        Currstge(icult, Joursim) = 1
    End If
    If (DVS(icult) >= TDV(icult, 1) And DVS(icult) < TDV(icult, 2)) Then
             Currstge(icult, Joursim) = 2
    ' prise en charge photopériode
             DL = dataclim.dDAYL(Joursim) + 0.9
             If (DL < MOPP(icult)) Then
                PPFAC = 1#
             Else
                PPFAC = 1# - (DL - MOPP(icult)) * SensPhot(icult)
             End If
             PPFAC = Min(1#, Max(0#, PPFAC))
             DVR(icult) = CTstade(icult, 2) * PPFAC
    End If
    If (DVS(icult) >= TDV(icult, 2) And DVS(icult) < TDV(icult, 3)) Then
        DVR(icult) = CTstade(icult, 3)
        Currstge(icult, Joursim) = 3
    End If
    If (DVS(icult) >= TDV(icult, 3) And DVS(icult) < TDV(icult, 4)) Then
        DVR(icult) = CTstade(icult, 4)
        Currstge(icult, Joursim) = 4
    End If
    If (DVS(icult) >= TDV(icult, 4) And DVS(icult) < TDV(icult, 5)) Then
        DVR(icult) = CTstade(icult, 5)
        Currstge(icult, Joursim) = 5
    End If
    If DVS(icult) >= TDV(icult, 5) Then
        Die(icult) = True
        DVR(icult) = 0
        Currstge(icult, Joursim) = 6
        If Currstge(icult, Joursim - 1) = 5 Then
            JourMat(icult) = Joursim
            Die(icult) = True
        End If

    End If
    '!-----CROPSTA: 0=before sowing; 1=day of sowing; 2=in seedbed;
    '!                  3=day of transplanting; 4=main growth period

    'attention dans ce qui suit voir s'il faut calculer TSTR par icult...
    If (cropsta(icult, Joursim) = 4) And cropsta(icult, Joursim - 1) = 3 Then TSTR = TS(icult)

    If (cropsta(icult, Joursim - 1) > 3 And TS(icult) < (TSTR + StrsChoc(icult))) Then DVR(icult) = 0#

    'attention StrsChoc doit etre nul si pas de repiquage
    'la date julpheno(icult, j) est la date à laquelle un nouveau currstge(icult) démarre julpheno1 le lendemain du dernier jour de currstge1
    If Currstge(icult, Joursim) > Currstge(icult, Joursim - 1) Then JulPheno(icult, Currstge(icult, Joursim)) = dataclim.nDOY(Joursim)

    End Sub
    Public Sub stressAzote(icult As Object, Joursim As Object, Navail As Double)
    'proposition du calcul du stress azoté inspiré de FIELD (QUEFTS) mais sur pas de temps journalier
    'pas testé de manière approfondie au 8/11/2017

    'On considère que l'efficience de capture de l'azote est de 1 une fois les pertes gazeuses
    'et par drainage retranchées à Navail, et considérant les autres nutriments non limitant
    'Nuptake = N availability
    'Mais dans l'absolu, dans FIELD (QUEFTS) NCtE = Nuptake/Navail

    NavailCult(icult, Joursim) = Navail + Nsymbjour(icult)

    If Biom(icult, Joursim - 1) > 0 Then
        NUPTtarget(icult, Joursim) = Biom(icult, Joursim - 1) / (NCvEmin(icult) + ((NCvEmax(icult) - NCvEmin(icult)) * alphaN(icult)))
        NRF(icult, Joursim) = NavailCult(icult, Joursim) / NUPTtarget(icult, Joursim)
            If NRF(icult, Joursim) > 1 Then NRF(icult, Joursim) = 1
        Else
        NRF(icult, Joursim) = 1
    End If
    'd
    'si alphaN = 0.5    NCvE = médiane (NCvEmin et NCvEmax)
    'si alphaN = 1      NCvE = NCvEmax
    'si alphaN = 0      NCvE = NCvEmin

    End Sub
    Public Sub Calcule_LAI(Vlaimax As Double, Joursim As Integer, icult As Integer, ContrainteW As Double, ActiveStressH As Boolean, ActivestressN As Boolean, LaiMC As Double, Nbcult As Double)

    Dim Ulai As Double
    Dim dLAI As Double
    Dim dLAISen As Double
    Dim TurfacN As Double

    ' constante suivante à transformer en variable !
    Const LAIseuilComp As Double = 0.1 'lai seuil à partir duquel la compétition pour la lumière a lieu

    If ContrainteW > (1 - SeuilTurg(icult)) Or Not (ActiveStressH) Then TurfacH(Joursim) = 1 Else TurfacH(Joursim) = ContrainteW / (1 - SeuilTurg(icult))

    'modif FA le 28/11/2019 introduction d'un effet moindre du stress N sur LAI que sur biomasse
    'If NRF(icult, joursim) = 1 Or Not (ActivestressN) Then TurfacN = 1 Else TurfacN = NRF(icult, joursim)
    If NRF(icult, Joursim) = 1 Or Not (ActivestressN) Then TurfacN = 1 Else TurfacN = NRF_lai(icult, Joursim)
    'fin modiFA

    Turfac(Joursim) = Min(TurfacH(Joursim), TurfacN)

    'est ce qu'il y a compétition pour la lumière entre les plantes de l'association :
    If Nbcult > 1 And LaiMC >= LAIseuilComp And cropsta(1, Joursim - 1) > 2 And cropsta(2, Joursim - 1) > 2 Then
        Competition(Joursim) = True
        Else
        Competition(Joursim) = False
    End If
    ' FA ci dessous prise en compte d ela compétition pour la lumière, reste à vérifier soigneusement en pas à pas...car routine executée pour chaque icult et là on affecte des valeurs pour les deux icults à chaque passage
    ' il aurait sans doute mieux valu introduire une routine spécifique au niveau culture...
    'calcul des facteurs de réduction du LAI liés à la compétition avec autre plante de l'association
    'pour chaque plante de l'association en fonction du rapport de LAI entre les deux au jour précédent
    If (Competition(Joursim)) Then
        If Lai(1, Joursim - 1) > Lai(2, Joursim - 1) Then
        CompFac(2, Joursim) = Exp(-CoefExtin(1) * Lai(1, Joursim - 1))
        CompFac(1, Joursim) = 1
        End If
        If Lai(1, Joursim - 1) < Lai(2, Joursim - 1) Then
        CompFac(1, Joursim) = Exp(-CoefExtin(2) * Lai(2, Joursim - 1))
        CompFac(2, Joursim) = 1
        End If
        Else
        CompFac(icult, Joursim) = 1
    End If

    If Transplant(icult) Then
        If cropsta(icult, Joursim) = 3 Then
            Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / TDV(icult, 1)

            dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
            dLAI = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim))
            Lai(icult, Joursim) = dLAI + Lai(icult, Joursim - 1)

        End If
    ' les test et cas peuvent être plus concis ?
    ' améliorer généricité via utilisation TDV
        If cropsta(icult, Joursim - 1) = 3 And cropsta(icult, Joursim) = 4 Then

            If Currstge(icult, Joursim) = 1 Then Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / 0.4

            If Currstge(icult, Joursim) = 2 Then Ulai = Vlaimax + (3 - Vlaimax) * (DVSt(icult, Joursim) - 0.4) / 0.25

            dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
            dLAI = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim)) * Turfac(Joursim) * CompFac(icult, Joursim)
    ' ligne spécifique du jour de repiquage. Cas du démariage à réfléchir
            Lai(icult, Joursim) = dLAI + (Lai(icult, Joursim - 1) * densplt(icult, Joursim) / densplt(icult, Joursim - 1))
        End If
    End If
    If cropsta(icult, Joursim - 1) > 3 Or Not Transplant(icult) Then

        If Currstge(icult, Joursim) = 1 Then Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / 0.4

        If Currstge(icult, Joursim) = 2 Then Ulai = Vlaimax + (3 - Vlaimax) * (DVSt(icult, Joursim) - 0.4) / 0.25

        dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
        dLAI = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim)) * Turfac(Joursim) * CompFac(icult, Joursim)
        Lai(icult, Joursim) = dLAI + Lai(icult, Joursim - 1)
        If Currstge(icult, Joursim) = 3 And Currstge(icult, Joursim - 1) = 2 Then JourLaiMax(icult) = Joursim - 1
        If Currstge(icult, Joursim) = 3 Or Currstge(icult, Joursim) = 4 Then Lai(icult, Joursim) = Lai(icult, Joursim - 1)
        If Currstge(icult, Joursim) = 5 And Currstge(icult, Joursim - 1) = 4 Then
        JourSen(icult) = Joursim
        Lai(icult, JourSen(icult)) = Lai(icult, Joursim - 1)

    End If
    ' introduction de l'effet de turfac sur LAI apres stade 5 (début sénéscence)
        If Currstge(icult, Joursim) = 5 Then
    '
    'dLAIsen moyen de la période de sénéscence des feuilles (négatif)
    ' attention  au cas où le LAI(JourSen) est > LAIrec...ne marche pas bien avec LAIRec <> 0 !!!!! et il faut Lairec <5 avec instruction ci-dessous
    If LAIrec(icult) > Lai(icult, JourSen(icult)) Then LAIrec(icult) = LAIrec(icult) * (Lai(icult, JourSen(icult)) / 5)
        dLAISen = (LAIrec(icult) - Lai(icult, JourSen(icult))) / (2 - DVSt(icult, JourSen(icult)))

        dLAISen = dLAISen * (DVSt(icult, Joursim) - DVSt(icult, Joursim - 1)) * (1 + SensiSen(icult) * (1 - Turfac(Joursim)))
     'si on veut pas d'effet du stress apres floraison: Sensisen=0
     ' effet maximal du stress apres floraison : sensisen=1: dLaisen est doublé, multiplication par deux de la vitesse de sénéscence
    ' Si Turfac vaut 1 (pas de stress), dLAI sen est inchangé
    ' A TESTER: prendre TurfacH seulement pour cet effet ?

        Lai(icult, Joursim) = Lai(icult, Joursim - 1) + dLAISen
        If Lai(icult, Joursim) < LAIrec(icult) Then Lai(icult, Joursim) = LAIrec(icult)


        End If

    If Currstge(icult, Joursim) = 6 Then Lai(icult, Joursim) = LAIrec(icult)
    End If

    End Sub
    Public Sub Calcule_LAI_SemiAride(Vlaimax As Double, Joursim As Integer, icult As Integer, ContrainteW As Double, ActiveStressH As Boolean, ActivestressN As Boolean, LaiMC As Double, Nbcult As Double)
    'variante
    ' pour introduire effet des stress sur la réduction du LAi dès après le stade LAI mw (apres Julpheno3)
    ' attention pas de LAIrecmax dans cette variante

    Dim Ulai As Double
    Dim dLAI As Double
    Dim dLAIpot As Double
    ' dlaipot et LAIpot(icult) ont été introduits mais ne sont pas variables explicatives ni envoyées vers table de sortie
    ' peut être pratique pour caler dlaimax sur un LAIpot -cible
    Dim dLAISen As Double
    Dim TurfacN As Double

    ' constante suivante à transformer en variable !
    Const LAIseuilComp As Double = 0.1 'lai seuil à partir duquel la compétition pour la lumière a lieu

    If ContrainteW > (1 - SeuilTurg(icult)) Or Not (ActiveStressH) Then TurfacH(Joursim) = 1 Else TurfacH(Joursim) = ContrainteW / (1 - SeuilTurg(icult))

    'modif FA le 28/11/2019 introduction d'un effet moindre du stress N sur LAI que sur biomasse
    'If NRF(icult, joursim) = 1 Or Not (ActivestressN) Then TurfacN = 1 Else TurfacN = NRF(icult, joursim)
    If NRF(icult, Joursim) = 1 Or Not (ActivestressN) Then TurfacN = 1 Else TurfacN = NRF_lai(icult, Joursim)
    'fin modiFA

    Turfac(Joursim) = Min(TurfacH(Joursim), TurfacN)

    'est ce qu'il y a compétition pour la lumière entre les plantes de l'association :
    If Nbcult > 1 And LaiMC >= LAIseuilComp And cropsta(1, Joursim - 1) > 2 And cropsta(2, Joursim - 1) > 2 Then
        Competition(Joursim) = True
        Else
        Competition(Joursim) = False
    End If
    ' FA ci dessous prise en compte de la compétition pour la lumière si culture associée, reste à vérifier soigneusement en pas à pas...car routine executée pour chaque icult et là on affecte des valeurs pour les deux icults à chaque passage
    ' il aurait sans ndout mieux valu introduire une routine spécifique au niveau culture...
    'calcul des facteurs de réduction du LAI liés à la compétition avec autre plante de l'association
    'pour chaque plante de l'association en fonction du rapport de LAI entre les deux au jour précédent
    If (Competition(Joursim)) Then
        If Lai(1, Joursim - 1) > Lai(2, Joursim - 1) Then
        CompFac(2, Joursim) = Exp(-CoefExtin(1) * Lai(1, Joursim - 1))
        CompFac(1, Joursim) = 1
        End If
        If Lai(1, Joursim - 1) < Lai(2, Joursim - 1) Then
        CompFac(1, Joursim) = Exp(-CoefExtin(2) * Lai(2, Joursim - 1))
        CompFac(2, Joursim) = 1
        End If
        Else
        CompFac(icult, Joursim) = 1
    End If

    If Transplant(icult) Then
        If cropsta(icult, Joursim) = 3 Then
            Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / TDV(icult, 1)

            dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
            dLAI = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim))
            dLAIpot = dLAI
            LAIpot(icult) = LAIpot(icult) + dLAIpot
            Lai(icult, Joursim) = dLAI + Lai(icult, Joursim - 1)

        End If
    ' les test et cas peuvent être plus concis ?
    ' améliorer généricité via utilisation TDV
        If cropsta(icult, Joursim - 1) = 3 And cropsta(icult, Joursim) = 4 Then

            If Currstge(icult, Joursim) = 1 Then Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / 0.4

            If Currstge(icult, Joursim) = 2 Then Ulai = Vlaimax + (3 - Vlaimax) * (DVSt(icult, Joursim) - 0.4) / 0.25

            dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
            dLAIpot = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim))
            dLAI = dLAIpot * Turfac(Joursim) * CompFac(icult, Joursim)
    ' ligne spécifique du jour de repiquage. Cas du démariage à réfléchir
            Lai(icult, Joursim) = dLAI + (Lai(icult, Joursim - 1) * densplt(icult, Joursim) / densplt(icult, Joursim - 1))
            LAIpot(icult) = dLAIpot + (LAIpot(icult) * densplt(icult, Joursim) / densplt(icult, Joursim - 1))
        End If
    End If
    If cropsta(icult, Joursim - 1) > 3 Or Not Transplant(icult) Then

        If Currstge(icult, Joursim) <= 2 Then
            If Currstge(icult, Joursim) = 1 Then Ulai = 1 + (Vlaimax - 1) * DVSt(icult, Joursim) / 0.4

            If Currstge(icult, Joursim) = 2 Then Ulai = Vlaimax + (3 - Vlaimax) * (DVSt(icult, Joursim) - 0.4) / 0.25


            dLAI = DLAImax(icult) / (1 + Exp(5.5 * (Vlaimax - Ulai))) * Hu(icult)
            dLAIpot = dLAI * deltaidens(icult, Joursim, densplt(icult, Joursim))
            dLAI = dLAIpot * Turfac(Joursim) * CompFac(icult, Joursim)
            Lai(icult, Joursim) = dLAI + Lai(icult, Joursim - 1)
            LAIpot(icult) = LAIpot(icult) + dLAIpot
        Else
            If Currstge(icult, Joursim) = 3 And Currstge(icult, Joursim - 1) = 2 Then JourLaiMax(icult) = Joursim - 1
            dLAISen = 0
            If Currstge(icult, Joursim) = 5 Then
                If Currstge(icult, Joursim - 1) = 4 Then
                    JourSen(icult) = Joursim
                    Lai(icult, JourSen(icult)) = Lai(icult, Joursim - 1)
                End If
                If LAIrec(icult) > Lai(icult, Joursim) Then LAIrec(icult) = Lai(icult, Joursim)
                dLAISen = (LAIrec(icult) - Lai(icult, JourSen(icult))) / (2 - DVSt(icult, JourSen(icult)))
            End If
            ' si sensisen=1 et turfac= 0 (stress maxi), la perte de LAI due au stress est égale au LAI du jour précédent
            ' rien n'interdit de fixer sensisen au dessus de 1
            ' si sensisen=0 pas de perte de LAI due au stress
            ' attention ici c'est turfac qui a été utilisé...peut être substituer par turfacH seulement ?
    '        dLAISen = dLAISen - SensiSen(icult) * (1 - Turfac(Joursim)) * Lai(icult, Joursim - 1)
    ' test de l'option TurfacH aout2022
            dLAISen = dLAISen - SensiSen(icult) * (1 - TurfacH(Joursim)) * Lai(icult, Joursim - 1)
    'fin test
            dLAISen = dLAISen * (DVSt(icult, Joursim) - DVSt(icult, Joursim - 1))
            Lai(icult, Joursim) = Lai(icult, Joursim - 1) + dLAISen
            If Lai(icult, Joursim) <= 0 Then Lai(icult, Joursim) = 0
            'If Currstge(icult, joursim) = 6 Then Lai(icult, joursim) = Max(LAIrec(icult), 0)
        End If

    End If


    End Sub
    Public Sub biomasse(dataclim As DataClimClass, ParSurRg As Double, icult As Integer, Joursim As Integer, ContrainteW As Double, ActiveStressH As Boolean, ActivestressN As Boolean)


    Dim WSfactN As Double
    Dim Ftemp As Double

    If ContrainteW > (1 - SeuilWS(icult)) Or Not (ActiveStressH) Then WSfactH(Joursim) = 1 Else WSfactH(Joursim) = ContrainteW / (1 - SeuilWS(icult))
    'modif FA 29/11/2019 effet différencié de stress N sur LAI et biomasse
    'If NRF(icult, joursim) = 1 Or Not (ActivestressN) Then WSfactN = 1 Else WSfactN = NRF(icult, joursim)
    If NRF(icult, Joursim) = 1 Or Not (ActivestressN) Then WSfactN = 1 Else WSfactN = NRF_bio(icult, Joursim)
    ' A améliorer on ne devrait pas utiliser WSfact ici mais un Sressfactor
    WSfact(Joursim) = Min(WSfactH(Joursim), WSfactN)

    If Transplant(icult) And cropsta(icult, Joursim) = 4 And cropsta(icult, Joursim - 1) = 3 Then
    raint(icult, Joursim) = 0.95 * ParSurRg * dataclim.dRg(Joursim) * (1 - Exp(-CoefExtin(icult) * (Lai(icult, Joursim) * densplt(icult, Joursim) / densplt(icult, Joursim - 1))))
    End If

    If Transplant(icult) And cropsta(icult, Joursim - 1) > 3 Or Not Transplant(icult) Then
    ' a vérifier : ne semble pas correct pour la compétition pour la lumière entre cultures associees!!!
    raint(icult, Joursim) = 0.95 * ParSurRg * dataclim.dRg(Joursim) * (1 - Exp(-CoefExtin(icult) * Lai(icult, Joursim)))
    End If

    ' il faut gérer la cascade du rayonnement à travers les deux espèces...et il faut choisir l'ordre de passage des icult dans l'équation en fonction de l'ordre de dominance


    If dataclim.dTmoy(Joursim) <= tcopt(icult) Then
        Ftemp = 1 - ((dataclim.dTmoy(Joursim) - tcopt(icult)) / (tcmin(icult) - tcopt(icult))) ^ 2
    Else
        Ftemp = 1 - ((dataclim.dTmoy(Joursim) - tcopt(icult)) / (tcmax(icult) - tcopt(icult))) ^ 2
    End If
    If Ftemp < 0 Then Ftemp = 0

    deltaBiom(icult, Joursim) = CO2fact(icult) * WSfact(Joursim) * PlantPReducFact * (Ebmax(icult) * raint(icult, Joursim) - 0.0815 * raint(icult, Joursim) ^ 2) * Ftemp / 100

    Biom(icult, Joursim) = deltaBiom(icult, Joursim) + Biom(icult, Joursim - 1)

    Nuptake(icult, Joursim) = deltaBiom(icult, Joursim) / (NCvEmin(icult) + ((NCvEmax(icult) - NCvEmin(icult)) * alphaN(icult)))
    SigmaNuptake(icult, Joursim) = SigmaNuptake(icult, Joursim - 1) + Nuptake(icult, Joursim)

    End Sub
    Public Sub Rendement(icult As Integer, Joursim As Integer, julsim As Integer)
    Dim ng As Integer
    'calcul du nombre de grains lorsqu'on atteint le stade 4 (début remplissage)
    If JulPheno(icult, 4) = julsim Then
        JourDrp(icult) = Joursim
        For ng = JourDrp(icult) - Nbjgrain(icult) + 1 To JourDrp(icult)
            Vitmoy(icult) = Vitmoy(icult) + deltaBiom(icult, ng)
        Next ng
        Vitmoy(icult) = 100 * Vitmoy(icult) / Nbjgrain(icult)
        ' vitmoy en g/m2/jour)
        Ngrains(icult) = Int(Cgrain(icult) * Vitmoy(icult) + Cgrainv0(icult))

        If Ngrains(icult) < 0 Then Ngrains(icult) = 0

        ' pour limiter le nombre de grains par pied à une valeur max du cuiltivar
        If (Ngrains(icult) / densplt(icult, Joursim) > Ngrmax(icult)) Then Ngrains(icult) = Ngrmax(icult) * densplt(icult, Joursim)
    End If
    ' entre stade 4 (début remplissage) et stade fin évolution Indice de récolte (maturité, 6): indice de recolte croissant

    If Currstge(icult, Joursim) >= 4 And Currstge(icult, Joursim) <= 6 Then
        IR(icult) = Vitircarb(icult) * (Joursim - JourDrp(icult) + 1)
        IR(icult) = Min(IR(icult), IRmax(icult))
        Grain(icult, Joursim) = Min(Biom(icult, Joursim) * IR(icult), P1grainMax(icult) * Ngrains(icult) / 100)
        If Ngrains(icult) <> 0 Then
        P1grain(icult, Joursim) = 100 * Grain(icult, Joursim) / Ngrains(icult)
        Else
        P1grain(icult, Joursim) = -999.9
        End If
    End If
    End Sub
    Public Sub Croirac(icult As Integer, Joursim As Integer, ZoneHumSousRac As Double, ZracMC As Double)
    ' attetion plante repiquées: prendre en compte stress repiquage sur descente racines ?
    Dim deltarac As Double
    If Currstge(icult, Joursim) < 4 Then
    ' si la plante n'est pas celle dont les racines sont les plus profondes,
    ' on ne tient pas compte du front d'humectation pour limiter le front racinaire
    ' introduire un test aussi sur Strac non nul ?

        If (Zrac(icult, Joursim - 1) < ZracMC) Then
            ZoneHumSousRac = ZoneHumSousRac + ZracMC - Zrac(icult, Joursim - 1)
        End If
        deltarac = Min(ZoneHumSousRac, DeltaRacMax(icult) * Hu(icult))
    End If
    Zrac(icult, Joursim) = Min(Zrac(icult, Joursim - 1) + deltarac, Zracmax(icult))

    End Sub
    'introduit le 24/04/13, fonction Stics, page 56 bouquin Nadine 2008
    Public Function FCO2(CO2c As Integer, icult As Integer) As Double
    FCO2 = 2 - Exp(Log(2 - alphaCO2(icult)) * (CO2c - 350) / (600 - 350))
    End Function
    Public Function deltaidens(icult As Integer, Joursim As Integer, densite As Double)
    deltaidens = densite
    If Lai(icult, Joursim - 1) > Laicomp(icult) Then
        If densite >= bdens(icult) Then
            deltaidens = deltaidens * (densite / bdens(icult)) ^ adens(icult)
        End If
    End If
    End Function
    Public Property ncropsta(icult As Integer, joursim As Integer) As Integer
        Get
            Return cropsta(icult, joursim)
        End Get
        Set(value As Integer)
            cropsta(icult, joursim) = value
        End Set
    End Property



    Public Property bDie(icult As Integer) As Boolean
        Get
            Return Die(icult)
        End Get
        Set(value As Boolean)
            Die(icult) = value
        End Set
    End Property



    Public Function nJulPheno(icult As Integer, Istade As Integer) As Integer
    Return JulPheno(icult, Istade)
    End Function
    Public Function dDVSt(icult As Integer, Joursim As Integer) As Double
    Return DVSt(icult, Joursim)
    End Function
    Public Function nCurrstge(icult As Integer, Joursim As Integer) As Integer
    Return Currstge(icult, Joursim)
    End Function
    Public Function sCultivar(icult As Integer) As String
    Return Cultivar(icult)
    End Function
    Public Function dSommeT(icult As Integer, Joursim As Integer) As Double
    Return SommeT(icult, Joursim)
    End Function
    Public Function dLAI(icult As Integer, Joursim As Integer) As Double
    Return Lai(icult, Joursim)
    End Function
    Public Function dBiom(icult As Integer, Joursim As Integer) As Double
    Return Biom(icult, Joursim)
    End Function
    Public Function nDeathDay(icult As Integer) As Integer
    Return Death_day(icult)
    End Function
    Public Function nJourMat(icult As Integer) As Integer
    Return JourMat(icult)
    End Function
    Public Function nJourSen(icult As Integer) As Integer
    Return JourSen(icult)
    End Function
    Public Function dGrain(icult As Integer, Joursim As Integer) As Double
    Return Grain(icult, Joursim)
    End Function
    Public Function dCoefExtin(icult As Integer) As Double
    Return CoefExtin(icult)
    End Function
    Public Function dZrac(icult As Integer, Joursim As Integer) As Double
    Return Zrac(icult, Joursim)
    End Function
    Public Function nZracmax(icult As Integer) As Integer
    Return Zracmax(icult)
    End Function
    Public Function nNgrains(icult As Integer) As Long
    Return Ngrains(icult)
    End Function
    Public Function dP1grain(icult As Integer, Joursim As Integer) As Double
    Return P1grain(icult, Joursim)
    End Function
    Public Function dVitmoy(icult As Integer) As Double
    Return Vitmoy(icult)
    End Function
    Public Function dKmax(icult As Integer) As Double
    Return Kmax(icult)
    End Function
    Public Function nLevSim(icult As Integer) As Integer
    Return LevSim(icult)
    End Function
    Public Function nZgraine(icult As Integer) As Integer
    Return Zgraine(icult)
    End Function
    Public Function dNsymb(icult As Integer) As Double
    Return Nsymb(icult)
    End Function
    Public Function dNavailCult(icult As Integer, Joursim As Integer) As Double
    Return NavailCult(icult, Joursim)
    End Function
    Public Function dNUPTtarget(icult As Integer, Joursim As Integer) As Double
    Return NUPTtarget(icult, Joursim)
    End Function
    Public Function dNRF(icult As Integer, Joursim As Integer) As Double
    Return NRF(icult, Joursim)
    End Function
    Public Function dNRF_lai(icult As Integer, Joursim As Integer) As Double
    Return NRF_lai(icult, Joursim)
    End Function
    Public Function dNRF_bio(icult As Integer, Joursim As Integer) As Double
    Return NRF_bio(icult, Joursim)
    End Function
    Public Function dSigmaNuptake(icult As Integer, Joursim As Integer) As Double
    Return SigmaNuptake(icult, Joursim)
    End Function
    Public Function dWSfactH(Joursim As Integer) As Double
    Return WSfactH(Joursim)
    End Function
    Public Function dWSfact(Joursim As Integer) As Double
    Return WSfact(Joursim)
    End Function
    Public Function dTurfacH(Joursim As Integer) As Double
    Return TurfacH(Joursim)
    End Function
    Public Function dTurfac(Joursim As Integer) As Double
    Return Turfac(Joursim)
    End Function
    Public Function dNuptake(icult As Integer, Joursim As Integer) As Double
    Return Nuptake(icult, Joursim)
    End Function
    Public Function dCompFac(icult As Integer, Joursim As Integer) As Double
    Return CompFac(icult, Joursim)
    End Function
    Public Function bCompetition(Joursim As Integer) As Boolean
    Return Competition(Joursim)
    End Function
    Public Function draint(icult As Integer, Joursim As Integer) As Double
    Return raint(icult, Joursim)
    End Function
    Public Function nJourLaiMax(icult As Integer) As Integer
    Return JourLaiMax(icult)
    End Function
    Public Function nDurCycMax(icult As Integer) As Integer
    Return DurCycMax(icult)
    End Function
    Public Function dNplant(icult As Integer) As Double
    Return Nplant(icult)
    End Function
    Public Function dRootABGRatio(icult As Integer) As Double
    Return RootABGRatio(icult)
    End Function
    Public Function dRootCN(icult As Integer) As Double
    Return RootCN(icult)
    End Function
    Public Function dNRootABGRatio(icult As Integer) As Double
    Return NRootABGRatio(icult)
    End Function
    Public Function dNGrainABGRatio(icult As Integer) As Double
    Return NGrainABGRatio(icult)
    End Function
    Public Function dPfactor(icult As Integer) As Double
    Return Pfactor(icult)
    End Function
End Class
