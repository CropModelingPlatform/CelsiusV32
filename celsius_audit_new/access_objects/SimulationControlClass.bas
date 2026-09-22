Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SimUnit As SimulationUnitClass
Dim IdentSimul As String
Dim MemoSU As SimulationUnitClass
Dim dataclim As DataClimClass
Dim dataIrrig As IrrigClass
Dim dataFertiMin As FertiMinClass
Dim dataFertiOrg As FertiOrgaClass
Dim GestionTechnique As GestionTechniqueClass
Dim plante As PLanteClass
Dim Culture As CultureClass
Dim GenParam As GenParamClass
Dim Sol As SolClass
Dim mulch As MulchClass
Dim OptionsModel As OptionsModelClass
Dim ApportOrga As ApportsOrgaClass

Public EtatInitial As EtatInitialClass
Public EtatFinal As EtatFinalClass
'VERSION 3 du 8-10 nov 2017'
'Modif FA du 30/08/18 pour scenarios CC:déplacement de la lecture des options avant lecture du climat..pas fini, pas opérationnel, pas gênant masi champs supplementaires lus
' modif FA du 7/09/18: cumul evap sur culture avec intro de la variable "pousse" dans Simulation
'modif FA nov. 2021 appel de la nouvelle procédure Sol.Sol.SeasonMinNC et appel adapté de la procedure modifiée Plante.StressAzoteOld
'appel de la procedure calculant TMS
' FA jan 22 introduction de "pousse" comme argument de Eausol pour calcul drainage sous profmax pendant période de culture
'FA aout 22 modif appel bilanfin pour tenir compte de Zracmax
' verifiée et nettoyée (commentaires et codes obsolètes) FA le 07/05/2026

Sub ReadParameters(rstSimulation As ADODB.Recordset, DataBase_Cnn As ADODB.Connection, comptesim As Long)
' A faire : introduire une mémoire de SimUnit pour comparer d'une ligne de SimUnitList à
' l'autre et économiser les lectures de tables qui ne seraient pas nécessaires
' selon modele fait pour GenParam (important surtout pour DataClim)
Dim icult As Integer



If comptesim > 0 Then Set MemoSU = SimUnit Else Set MemoSU = New SimulationUnitClass
Set SimUnit = New SimulationUnitClass
Call SimUnit.ReadSimUnitParameters(rstSimulation)
IdentSimul = SimUnit.sIdSim
If comptesim = 0 Or (SimUnit.sidGenParam <> MemoSU.sidGenParam) Then
    Set GenParam = New GenParamClass
    Call GenParam.LisGenParam(DataBase_Cnn, SimUnit)
End If
If comptesim = 0 Or (SimUnit.nidCodModel <> MemoSU.nidCodModel) Then
    Set OptionsModel = New OptionsModelClass
    Call OptionsModel.LisOptionsModel(DataBase_Cnn, SimUnit.nidCodModel)
End If

Set dataclim = New DataClimClass
Call dataclim.LisClimD(DataBase_Cnn, SimUnit)

Set EtatInitial = New EtatInitialClass
If Codesuite = 0 Then
    Call EtatInitial.LisInitialData(DataBase_Cnn, SimUnit.sidIni)
Else
    Call EtatInitial.recursive(EtatFinal)
End If

Set GestionTechnique = New GestionTechniqueClass
Call GestionTechnique.LisTech(DataBase_Cnn, SimUnit)
 Set dataIrrig = New IrrigClass
If GestionTechnique.bIrrigON Then Call dataIrrig.LisIrrD(DataBase_Cnn, SimUnit, GestionTechnique)

 Set dataFertiMin = New FertiMinClass
If GestionTechnique.bfertiminON Then Call dataFertiMin.LisFertiMinD(DataBase_Cnn, SimUnit, GestionTechnique)
  Set dataFertiOrg = New FertiOrgaClass
If GestionTechnique.bfertiorgON Then
    Call dataFertiOrg.LisFertiOrgD(DataBase_Cnn, SimUnit, GestionTechnique)
    If dataFertiOrg.sTypeMorga <> "indetermine" Then
      Set ApportOrga = New ApportsOrgaClass
      Call ApportOrga.LisApportsOrga(dataFertiOrg.sTypeMorga)
    End If
End If
Call dataclim.update_DAP(SimUnit.nStartDay, GestionTechnique.niplt(1))


If OptionsModel.bCorAlti Then Call dataclim.CorrigTAlt(GestionTechnique.nAltiCult, SimUnit.nNbJourSimul)
Set plante = New PLanteClass
Call plante.LisPlante(DataBase_Cnn, GestionTechnique)
Call dataclim.TempMoySaison(Max(plante.nDurCycMax(1), plante.nDurCycMax(2)), SimUnit.nNbJourSimul)
Set Sol = New SolClass
Call Sol.LisSol(DataBase_Cnn, SimUnit, plante.nZgraine(1), GestionTechnique.bDriveRuiObs)

Set Culture = New CultureClass
For icult = 1 To GestionTechnique.nNbCult
    Call plante.Iniplante(icult, GestionTechnique.dDensSem(icult), GestionTechnique.niplt(icult), GestionTechnique.dDensRepiqu(icult), GestionTechnique.nirepiqu(icult), SimUnit.nStartDay, SimUnit.nEndDOY, Min(Sol.iZtotsol, Sol.nZObstacleRac), dataclim.nCO2c, Sol.dPRedFact)
    Call Culture.InitMixedCanopy(plante.dCoefExtin(icult), plante.nZracmax(icult), plante.dKmax(icult), plante.dPfactor(icult))
    Next icult
If Codesuite = 0 Then
    Call Sol.initsol(Culture.dZracmaxMC, EtatInitial.dStockinit, EtatInitial.bIniSolhautON)
    Else
    Call Sol.InitSolRecurs(EtatFinal)
End If
Call Culture.lisCultureMC
Set mulch = New MulchClass
Call mulch.LisMulch(DataBase_Cnn, GestionTechnique.nCodParamMulch, EtatInitial.dQpaillisinit)
If GestionTechnique.bDriveRuiObs Then Call mulch.LisRuiObs(DataBase_Cnn, SimUnit)

'MsgBox (SimUnit.sIdSim)

End Sub
Sub Simulation()
Dim Joursim As Integer
Dim icult As Integer
Dim pousse As Boolean



If OptionsModel.sCodeDevelop = "Oryza" Then Call plante.AdapteCT(GestionTechnique.nNbCult)
If OptionsModel.nTypeNPKstress = 1 Then Call Sol.SeasonMinNC(GestionTechnique.dtApportMON, GestionTechnique.dCsurN_AO, GestionTechnique.dKresY, dataclim.dTMS, GestionTechnique.dtApportMinN + EtatInitial.dNinit + dataclim.dNprecipTot)

'debut boucle sur pas de temps
    For Joursim = 1 To SimUnit.nNbJourSimul
    
' à améliorer pour ne calculer que pour les plantes pas mortes
       For icult = 1 To GestionTechnique.nNbCult
' seul calcul stess N saisonnier possible ici avec deux modalités (relations linéraire entre stress et rapport "offre/demande" ou relation courbe
       If OptionsModel.bActiveStressN Then
' choix de versions lineaires et courbe ligne suivante selon indication dans OptionsModel (appel possible de la version journalière dans la boucle sur les jours si TypeNPKstress=2 )
          If OptionsModel.nTypeNPKstress = 0 Then Call plante.stressAzoteOld(icult, Joursim, Sol.dNmintot, GestionTechnique.dtApportMinN + EtatInitial.dNinit + dataclim.dNprecipTot)
          If OptionsModel.nTypeNPKstress = 1 Then Call plante.stressAzoteOldCourbe(icult, Joursim, Sol.dNmintot, GestionTechnique.dtApportMinN + EtatInitial.dNinit + dataclim.dNprecipTot)
       End If
       If OptionsModel.bCyberST Then Call GestionTechnique.CyberPlouck(icult, plante.bDie(icult), plante.nDeathDay(icult), Joursim, SimUnit.nStartDay)
       If GestionTechnique.bRessemis(icult) And Joursim > 2 Then
         Call GestionTechnique.SemisAuto(icult, Joursim, SimUnit.nStartDay, ((dataclim.dPlu(Joursim - 1) + dataIrrig.dIrr(Joursim - 1)) - mulch.dRuis(Joursim - 1)) + ((dataclim.dPlu(Joursim - 2) + dataIrrig.dIrr(Joursim - 2)) - mulch.dRuis(Joursim - 2)))
' attention si la date de début de la période de semis est inférieure à Startday+2, on ne semetra pas avant le 2ejour même si le 1er aurait satisfait la condition
       End If
       
    If Not GestionTechnique.bRessemis(icult) And Joursim = GestionTechnique.niplt(icult) - SimUnit.nStartDay + 1 Then
            Call plante.Iniplante(icult, GestionTechnique.dDensSem(icult), GestionTechnique.niplt(icult), GestionTechnique.dDensRepiqu(icult), GestionTechnique.nirepiqu(icult), SimUnit.nStartDay, SimUnit.nEndDOY, Min(Sol.iZtotsol, Sol.nZObstacleRac), dataclim.nCO2c, Sol.dPRedFact)
            ' à faire: voir si cette intialisation se substitue complètement à celle qui suit la lecture du fichier plante
            plante.bDie(icult) = False
            plante.ncropsta(icult, Joursim - 1) = 1
            Call GestionTechnique.CompteSemis(icult)
            
        End If
       Next icult
       

       If Not (plante.bDie(1) And plante.bDie(2)) Then
' s'il y a au moins une plante en croissance

'
               
                For icult = 1 To GestionTechnique.nNbCult
                If Not plante.bDie(icult) Then
                    pousse = True
                    plante.ncropsta(icult, Joursim) = plante.ncropsta(icult, Joursim - 1)
'                    If Plante.ncropsta(icult, joursim) = 3 And GestionTechnique.bSerreTunnelON Then Call dataClim.CorrigTSerres(joursim) 'appel du modele de temperature sous serre tunnel
                    If plante.ncropsta(icult, Joursim) >= 1 And plante.ncropsta(icult, Joursim) < 3 Then
                        Call plante.GerminLevee(dataclim.nDOY(Joursim), icult, Joursim, GestionTechnique.nilev(icult), OptionsModel.bSimLevee, Sol.bContrainteHlevee, dataclim.dTmoy(Joursim))
                    End If
                    If plante.ncropsta(icult, Joursim) >= 3 Then
                        If plante.ncropsta(icult, Joursim) < 4 Then
                            If dataclim.nDOY(Joursim) >= GestionTechnique.nirepiqu(icult) Or Not GestionTechnique.bRepiquageON(icult) Then plante.ncropsta(icult, Joursim) = 4
                        End If
' vérifier si positionnement optimal appel plante.stressAzote (version dérivée de Field, couplée pour l'instant à un bilan N journalier dérivé de Stics) ici.
                        If OptionsModel.bActiveStressN And OptionsModel.nTypeNPKstress = 2 Then Call plante.stressAzote(icult, Joursim, Sol.dNavail(Joursim - 1))
     'à faire ? repasser signal fletrissement en scalaire ?
                        If OptionsModel.sCodeDevelop = "Oryza" Then
                            Call plante.pheno_sigmaT(dataclim, Joursim, Sol.nSignalFletrissement(Joursim - 1), icult)
                        Else
                            ' codeDevelop est supposé etre "Direct"
                            Call plante.phenoCTphot(dataclim, Joursim, Sol.nSignalFletrissement(Joursim - 1), icult)
                        End If
' choix ici en dur dans le code de la méthode de calcul du LAI...à introduire comme une option ?
'                        Call Plante.Calcule_LAI(GenParam.dVlaimax, joursim, icult, Sol.dContrainteW, OptionsModel.bActiveStressH, OptionsModel.bActiveStressN, Culture.dLaiMC(joursim - 1), GestionTechnique.nNbCult)
                        Call plante.Calcule_LAI_SemiAride(GenParam.dVlaimax, Joursim, icult, Sol.dContrainteW, OptionsModel.bActiveStressH, OptionsModel.bActiveStressN, Culture.dLaiMC(Joursim - 1), GestionTechnique.nNbCult)
                        Call plante.biomasse(dataclim, GenParam.dParSurRg, icult, Joursim, Sol.dContrainteW, OptionsModel.bActiveStressH, OptionsModel.bActiveStressN)
                        Call plante.Rendement(icult, Joursim, dataclim.nDOY(Joursim))
                        Call plante.Croirac(icult, Joursim, Sol.dZoneHumSousRacines, Culture.dZRacineMC(Joursim - 1))
                        Call Culture.LaiMixedCanopy(Joursim, plante.dLAI(icult, Joursim))
                        Call Culture.CroiRacMixedCrop(Joursim, plante.dZrac(icult, Joursim))
                        Call Culture.NuptakeMixedCrop(Joursim, plante.dNuptake(icult, Joursim))
                    End If
                End If
                Next icult
                
     
        End If 'fin  condition sur présence plantes
    Call mulch.BiomasseMulch(Joursim, dataclim.nDOY(Joursim) = GestionTechnique.nimulch, GestionTechnique.dQpaillisApport)
    Call Culture.Evaporation_Pot_SolMulch(Joursim, dataclim.dEtp(Joursim))
    'on appelle bilanmulch même si pas de mulch
    Call mulch.Ruissellement(Joursim, dataclim.dPlu(Joursim) + dataIrrig.dIrr(Joursim), Sol.ClTypeSurf, Culture.dLaiMC(Joursim))
    Call mulch.BilanMulch(Joursim, Culture.dEoSM, mulch.dEau_vers_Mulch)
    Call Sol.evaporation(Joursim, Culture.dEoSM, mulch.dEomulch, pousse)
    'TODO faudrait faire d'abord actualisation des stocks avant transpi puis transpi puis réactualiser stocks...
    Call Culture.CalcTranspiMC(Joursim, dataclim.dEtp(Joursim), Sol.dEsol(Joursim), mulch.dEmulch(Joursim), Sol.dContrainteW)
           
    Call Sol.EauSol(Joursim, mulch.dEauVersSol, Culture.dTranspiMC(Joursim), Culture.dZRacineMC(Joursim), Culture.dZRacineMC(Joursim) - Culture.dZRacineMC(Joursim - 1), Culture.dZracmaxMC, pousse)
    'Attention ci-dessous option de bilan N journalier en chantier, verifications necessaires
    'pourquoi conditionner le bilan à l'activation du stress ?
    If OptionsModel.bActiveStressN And OptionsModel.nTypeNPKstress = 2 Then
        Call Sol.TempSoil(dataclim.dTmoy(Joursim))
        Call Sol.MinNorgSS(Joursim)
        Call Sol.DenitNavail(Joursim)
        'attention ligne suivante serait interessant de parametrer le CsurNhum, fixé à 12 en l'état ?
        Call Sol.MinNorgapporteStics(Joursim, dataFertiOrg.dQorga(Joursim), dataFertiOrg.dCsurNRes, ApportOrga)
        Call Sol.StockNavail(Joursim, dataFertiMin.dNmin(Joursim), Culture.dNuptMC(Joursim), dataclim.dPlu(Joursim) * dataclim.dConcNplu)
              
    End If
    Call Sol.ConcNEauSol(Joursim, Culture.dZracmaxMC)
    Next Joursim
If OptionsModel.nTypeNPKstress = 1 Then
    Call plante.ExportN(GestionTechnique.dtApportMinN + EtatInitial.dNinit + Sol.dNmintot + dataclim.dNprecipTot)
    Call Sol.BilanNFin(plante.dNplant(1) + plante.dNplant(2), GestionTechnique.dtApportMinN + EtatInitial.dNinit + dataclim.dNprecipTot, SimUnit.nNbJourSimul, Culture.dZracmaxMC)
End If
End Sub
Sub SortieSynthesis(TabSynt As ADODB.Recordset)
Dim Istade As Integer
Dim ErrFa As Boolean
Dim msg As String


On Error GoTo Err_SortieSynthesis
'update table "OutputSynt" with synthetis of the results



TabSynt.AddNew
TabSynt.Fields(0).Value = SimUnit.sIdSim
For Istade = 1 To 6
TabSynt.Fields(Istade).Value = plante.nJulPheno(1, Istade)
Next Istade
TabSynt.Fields(7).Value = plante.nDeathDay(1)
TabSynt.Fields(8).Value = plante.dBiom(1, plante.nJourMat(1))
TabSynt.Fields(9).Value = plante.dGrain(1, plante.nJourMat(1))
TabSynt.Fields(10).Value = plante.dLAI(1, plante.nJourLaiMax(1))
TabSynt.Fields(11).Value = Sol.dSigmaSimEsol
TabSynt.Fields(12).Value = Sol.dSigmaSimDr
TabSynt.Fields(13).Value = Sol.dSigmaSimDrprofmax
TabSynt.Fields(14).Value = Sol.dStockSol(0)
TabSynt.Fields(15).Value = Sol.dStockSol(SimUnit.nNbJourSimul)
TabSynt.Fields(16).Value = mulch.dSigmaSimEmulch
TabSynt.Fields(17).Value = mulch.dSigmaSimRuis
TabSynt.Fields(18).Value = Culture.dSigmaTranspiMC
TabSynt.Fields(19).Value = mulch.dSigmaSimPluM
TabSynt.Fields(20).Value = plante.nNgrains(1)
TabSynt.Fields(21).Value = plante.dP1grain(1, plante.nJourMat(1))
TabSynt.Fields(22).Value = plante.dVitmoy(1)
TabSynt.Fields(23).Value = GestionTechnique.niplt(1)
TabSynt.Fields(24).Value = GestionTechnique.nNbsemis(1)
For Istade = 1 To 6
TabSynt.Fields(24 + Istade).Value = plante.nJulPheno(2, Istade)
Next Istade
TabSynt.Fields(31).Value = plante.nDeathDay(2)
TabSynt.Fields(32).Value = plante.dBiom(2, plante.nJourMat(2))
TabSynt.Fields(33).Value = plante.dGrain(2, plante.nJourMat(2))
TabSynt.Fields(34).Value = plante.dLAI(2, plante.nJourSen(2))
TabSynt.Fields(35).Value = plante.nNgrains(2)
TabSynt.Fields(36).Value = plante.dP1grain(2, plante.nJourMat(2))
TabSynt.Fields(37).Value = plante.dVitmoy(2)
TabSynt.Fields(38).Value = GestionTechnique.niplt(2)
TabSynt.Fields(39).Value = GestionTechnique.nNbsemis(2)
TabSynt.Fields(40).Value = Sol.bBilanNnonOK
TabSynt.Fields(41).Value = Sol.dStockN
TabSynt.Fields(42).Value = mulch.bRuisEtrange
TabSynt.Fields(43).Value = Sol.dSigmaCultEsol
TabSynt.Fields(44).Value = Sol.dtxminN
TabSynt.Fields(45).Value = Sol.dFminY1
TabSynt.Fields(46).Value = Sol.dStockN1
TabSynt.Fields(47).Value = Sol.dStockC1
TabSynt.Fields(48).Value = Sol.dCmintot
TabSynt.Fields(49).Value = Sol.dCminRes
TabSynt.Fields(50).Value = Sol.dNminMOSY
TabSynt.Fields(51).Value = Sol.dNminRes
TabSynt.Fields(52).Value = Sol.dNreliquat
TabSynt.Fields(53).Value = plante.dNplant(1)
TabSynt.Fields(54).Value = plante.dNplant(2)
TabSynt.Fields(55).Value = plante.dNplant(1) * plante.dNGrainABGRatio(1)
TabSynt.Fields(56).Value = plante.dNplant(1) * plante.dNRootABGRatio(1)
TabSynt.Fields(57).Value = plante.dNplant(1) * plante.dNRootABGRatio(1) * plante.dRootCN(1)
TabSynt.Fields(58).Value = plante.dBiom(1, plante.nJourMat(1)) * plante.dRootABGRatio(1)
TabSynt.Fields(59).Value = Sol.dSigmaCultDrprofmax
TabSynt.Fields(60).Value = Sol.dCumPerteNdrainCult
TabSynt.Fields(61).Value = Sol.dNminMOSY * (plante.nJulPheno(1, 6) - GestionTechnique.niplt(1)) / SimUnit.nNbJourSimul
'ajout special Agmip
TabSynt.Fields(62).Value = Sol.dStsurf(GestionTechnique.niplt(1) - SimUnit.nStartDay + 1)
TabSynt.Fields(63).Value = Sol.dStockC
TabSynt.Fields(64).Value = dataclim.dNprecipTot
TabSynt.Fields(65).Value = Sol.dCumPerteNdrain
TabSynt.Fields(66).Value = EtatInitial.dNinit
TabSynt.Update

Exit_SortieSynthesis:
    Exit Sub
ErrFA_SortieSynthesis:
    ErrFa = True
    MsgBox ("ERREUR ! " & msg)
    Resume Exit_SortieSynthesis
Err_SortieSynthesis:
    MsgBox Err.Description
    Resume Exit_SortieSynthesis

End Sub
Sub EcritDresu(Db_Cnn As ADODB.Connection, comptesim As Long)

Dim jour As Integer
Dim icult As Integer
Dim TabDayPlante(1 To 2) As ADODB.Recordset
'TODO: integrer modifs proposées par Guillaume sur conseils IA (begintrans, CommitTrans et vidage table antérieure par SQL)

    For icult = 1 To GestionTechnique.nNbCult
        Set TabDayPlante(icult) = New ADODB.Recordset


        TabDayPlante(icult).Open "OutputD_" & Trim(Str(icult)), Db_Cnn, , adLockOptimistic
' vidage tables sorties si 1ere simul (serait plus rapide par SQL)
        If comptesim = 0 Then
            While Not TabDayPlante(icult).EOF
                TabDayPlante(icult).Delete
                TabDayPlante(icult).MoveNext
            Wend
        End If
    If OptionsModel.bEcritDresus Then
        If comptesim > 0 And Not TabDayPlante(icult).EOF Then TabDayPlante(icult).MoveLast
      
        For jour = 1 To SimUnit.nNbJourSimul
   
            TabDayPlante(icult).AddNew
            TabDayPlante(icult).Fields(0).Value = SimUnit.sIdSim & Str(dataclim.nDAP(jour))
             TabDayPlante(icult).Fields(1).Value = SimUnit.sIdSim
            TabDayPlante(icult).Fields(2).Value = SimUnit.sIdWeather
            TabDayPlante(icult).Fields(3).Value = dataclim.nCurrentYear(jour)
            TabDayPlante(icult).Fields(4).Value = SimUnit.sIdTec
            TabDayPlante(icult).Fields(5).Value = plante.sCultivar(icult)
            TabDayPlante(icult).Fields(6).Value = dataclim.nDAP(jour)
            TabDayPlante(icult).Fields(7).Value = dataclim.nDOY(jour)
            TabDayPlante(icult).Fields(8).Value = dataclim.dTmax(jour)
            TabDayPlante(icult).Fields(9).Value = dataclim.dTmin(jour)
            TabDayPlante(icult).Fields(10).Value = plante.dDVSt(icult, jour)
            TabDayPlante(icult).Fields(11).Value = plante.nCurrstge(icult, jour)
            TabDayPlante(icult).Fields(12).Value = plante.ncropsta(icult, jour)
            TabDayPlante(icult).Fields(13).Value = plante.dSommeT(icult, jour)
            TabDayPlante(icult).Fields(14).Value = plante.dLAI(icult, jour)
            TabDayPlante(icult).Fields(15).Value = plante.dBiom(icult, jour)
            TabDayPlante(icult).Fields(16).Value = plante.dGrain(icult, jour)
            TabDayPlante(icult).Fields(17).Value = plante.dZrac(icult, jour)
            TabDayPlante(icult).Fields(18).Value = Sol.dStsurf(jour)
            TabDayPlante(icult).Fields(19).Value = Sol.dEsol(jour)
            TabDayPlante(icult).Fields(20).Value = Sol.dStnonrac(jour)
            TabDayPlante(icult).Fields(21).Value = Sol.dStrac(jour)
            TabDayPlante(icult).Fields(22).Value = Culture.dTranspiMC(jour)
            TabDayPlante(icult).Fields(23).Value = Sol.dDrprofmax(jour)
            
            
            TabDayPlante(icult).Fields(24).Value = Sol.dStprofond(jour)
            TabDayPlante(icult).Fields(25).Value = Sol.dStockSol(jour)
            TabDayPlante(icult).Fields(26).Value = dataclim.dPlu(jour)
            TabDayPlante(icult).Fields(27).Value = Sol.dStockMes(jour)
            
            TabDayPlante(icult).Fields(28).Value = mulch.dEmulch(jour)
            TabDayPlante(icult).Fields(29).Value = mulch.dRuis(jour)
            TabDayPlante(icult).Fields(30).Value = mulch.dQpaillis(jour)
            TabDayPlante(icult).Fields(31).Value = plante.dP1grain(icult, jour)
           
            'TabDayPlante(icult).Fields(32).Value = dataclim.dEtp(jour)
            TabDayPlante(icult).Fields(32).Value = dataclim.dRg(jour)
            TabDayPlante(icult).Fields(33).Value = Culture.dTPotMC(jour)
            TabDayPlante(icult).Fields(34).Value = Sol.bDrainageON(jour)
            TabDayPlante(icult).Fields(35).Value = Sol.dConcNsol(jour)
            TabDayPlante(icult).Fields(36).Value = Sol.dDr(jour)
            TabDayPlante(icult).Fields(37).Value = Sol.dperteNDrain(jour)
            TabDayPlante(icult).Fields(38).Value = plante.dNuptake(icult, jour)
            TabDayPlante(icult).Fields(39).Value = plante.dSigmaNuptake(icult, jour)
            TabDayPlante(icult).Fields(40).Value = Sol.dNavail(jour) * 0.01
            TabDayPlante(icult).Fields(41).Value = dataFertiMin.dNmin(jour)
            TabDayPlante(icult).Fields(42).Value = dataFertiOrg.dNorg(jour)
            TabDayPlante(icult).Fields(43).Value = Sol.dNavail(jour)
            TabDayPlante(icult).Fields(44).Value = plante.dNUPTtarget(icult, jour)
            TabDayPlante(icult).Fields(45).Value = plante.dNRF(icult, jour)
            TabDayPlante(icult).Fields(46).Value = plante.dWSfactH(jour)
            TabDayPlante(icult).Fields(47).Value = plante.dWSfact(jour)
            TabDayPlante(icult).Fields(48).Value = plante.dTurfacH(jour)
            TabDayPlante(icult).Fields(49).Value = plante.dTurfac(jour)
            TabDayPlante(icult).Fields(50).Value = Sol.dStger(jour)
            TabDayPlante(icult).Fields(51).Value = plante.bCompetition(jour)
            TabDayPlante(icult).Fields(52).Value = plante.dCompFac(icult, jour)
            TabDayPlante(icult).Fields(53).Value = plante.draint(icult, jour)
            TabDayPlante(icult).Fields(54).Value = Sol.dStCJ(jour)
            TabDayPlante(icult).Fields(55).Value = Sol.dStNJ(jour)
            TabDayPlante(icult).Fields(56).Value = 3.67 * Sol.dCminMOSJour(jour)
            TabDayPlante(icult).Fields(57).Value = 3.67 * Sol.dCminResJour(jour)
            TabDayPlante(icult).Fields(58).Value = Sol.dNmin(jour)
            TabDayPlante(icult).Fields(59).Value = Sol.dMinNorgInput(jour)
            
            'TabDayPlante(icult).Fields(54).Value = Sol.dStockC + (Sol.dStockC1 - Sol.dStockC) * jour / SimUnit.nNbJourSimul
            'TabDayPlante(icult).Fields(55).Value = Sol.dStockN + (Sol.dStockN1 - Sol.dStockN) * jour / SimUnit.nNbJourSimul
            'TabDayPlante(icult).Fields(56).Value = 3.67 * (Sol.dCmintot - Sol.dCminRes) / SimUnit.nNbJourSimul
            'TabDayPlante(icult).Fields(57).Value = 3.67 * Sol.dCminRes / SimUnit.nNbJourSimul
            'TabDayPlante(icult).Fields(58).Value = Sol.dNminMOSY / SimUnit.nNbJourSimul
            'TabDayPlante(icult).Fields(59).Value = Sol.dNminRes / SimUnit.nNbJourSimul
            TabDayPlante(icult).Update

           
        Next jour
    End If 'fin cas où ecritdresus est vrai
        TabDayPlante(icult).Close
        Set TabDayPlante(icult) = Nothing
    Next icult
End Sub
Sub MemoEtatFinal()
Set EtatFinal = New EtatFinalClass
Call EtatFinal.EcritEF2(Sol, plante, mulch, SimUnit.nNbJourSimul)
End Sub
Property Get Coptionsmodel() As OptionsModelClass
Coptionsmodel = OptionsModel
End Property
Property Get sIdentSimul() As String
sIdentSimul = IdentSimul
End Property