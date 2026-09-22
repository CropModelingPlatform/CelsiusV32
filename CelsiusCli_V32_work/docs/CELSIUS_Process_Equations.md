# CELSIUS V32 Process Equations

This document rewrites the main CELSIUS V32 algorithms as mathematical process descriptions inferred from the code. It is intended as a model-oriented companion to [CELSIUS_Model_Overview.md](CELSIUS_Model_Overview.md).

> **V32 revision status (September 2026).** This revision is tied to the
> generated VB.NET sources in `CelsiusCli_V32_work/Converted`, which are
> reproduced exactly from `celsius_audit_new/access_objects`. The formulations
> remain implementation-derived and have not undergone an independent
> equation-by-equation scientific validation.

It is not an official scientific specification. It is an interpretation of the current implementation, mainly from:

- [SimulationControlClass.vb](../Converted/SimulationControlClass.vb)
- [PLanteClass.vb](../Converted/PLanteClass.vb)
- [CultureClass.vb](../Converted/CultureClass.vb)
- [MulchClass.vb](../Converted/MulchClass.vb)
- [SolClass.vb](../Converted/SolClass.vb)
- [DataClimClass.vb](../Converted/DataClimClass.vb)
- [Functions.vb](../Converted/Functions.vb)

## 1. Scope and conventions

CELSIUS is a daily time-step crop-soil-water-mulch model with optional nitrogen stress. The implementation is written for up to two crops in mixture.

Unless otherwise stated:

- `j` is the simulation day index
- `i` is the crop index
- daily forcing comes from `DataClimClass`
- most water state variables are expressed in `mm`
- canopy and crop variables are often in `t/ha`, `kg/ha`, or adimensional form depending on the variable

The state update order matters. The equations below follow the actual execution order in `SimulationControlClass.Simulation()`.

## 2. Global structure of one simulation

For each simulation unit:

1. read static parameters and forcing series
2. initialize plant, canopy, soil, and mulch states
3. for each day:
   - manage sowing and reseeding
   - update crop development and growth
   - update mulch and runoff
   - update soil evaporation and water balance
   - update canopy transpiration
   - optionally update daily nitrogen balance
4. write summary and daily outputs

In compact form:

```text
Inputs -> Initialization -> Daily loop -> Outputs
```

## 3. Climate preprocessing

### 3.1 Mean air temperature

From [DataClimClass.vb](../Converted/DataClimClass.vb):

```text
Tmoy_j = (Tmin_j + Tmax_j) / 2
```

### 3.2 Astronomical day length

The model computes day length with:

```text
lat_rad = pi * lat / 180
delta = 0.409 * sin(2*pi*J/365 - 1.39)
omega = -atan(-tan(lat_rad)*tan(delta) / sqrt(1 - tan(lat_rad)^2 * tan(delta)^2)) + 2*atan(1)
DAYL_j = 24 * omega / pi
```

This is used in photoperiod-sensitive phenology.

### 3.3 Optional altitude correction

If altitude correction is enabled:

```text
Tmin'_j = Tmin_j - 0.6 * (AltiCult - AltitudeStation) / 100
Tmax'_j = Tmax_j - 0.6 * (AltiCult - AltitudeStation) / 100
Tmoy'_j = (Tmin'_j + Tmax'_j) / 2
```

## 4. Simulation calendar and identity

Each row of `SimUnitList` defines:

- simulation identifier
- weather series
- soil
- technical management
- initial state
- general parameters
- model options
- simulation start and end dates

The length of simulation is:

```text
EndDOY = EndDay + ((EndYear - StartYear) mod 2) * (365 - I_bissextile)
NbJourSimul = EndDOY - StartDay + 1
```

where `I_bissextile = 1` for leap years and `0` otherwise.

## 5. Crop establishment and emergence

### 5.1 Thermal time for emergence

From [PLanteClass.vb](../Converted/PLanteClass.vb):

```text
HUleve_j = max(Tmoy_j - Tger_i, 0)
```

When emergence is simulated:

```text
TSlevee_i(j) = TSlevee_i(j-1) + HUleve_j
```

Emergence occurs when:

```text
TSlevee_i(j) >= CTlevee_i
```

### 5.2 Emergence water constraint

Emergence is blocked by the germination layer water status. The germination stock is `Stger`.

The code uses hard thresholds:

```text
if Stger_j >= 0.14 * TAW * Zger then ContrainteHlevee = False
if Stger_j <= 0.10 * TAW * Zger then ContrainteHlevee = True
```

This is not a smooth stress response. It is a Boolean gate.

## 6. Thermal time and phenology

CELSIUS contains two phenology schemes.

### 6.1 Daily heat units

The main heat unit function is:

```text
HU_j = max(min(Tmoy_j, tdmax_i) - tdmin_i, 0)
```

### 6.2 Direct cumulative thermal-time phenology

In `phenoCTphot`, cumulative thermal time is:

```text
SommeT_i(j) = SommeT_i(j-1) + HU_j
```

With photoperiod effect during stage 2 only:

```text
DL_j = DAYL_j + 0.9

if DL_j < MOPP_i:
    PPFAC_j = 1
else:
    PPFAC_j = 1 - (DL_j - MOPP_i) * SensPhot_i

PPFAC_j = min(1, max(0, PPFAC_j))
```

Developmental thermal time is then:

```text
TxDev_j = HU_j * PPFAC_j
```

with transplant shock gate:

```text
StopTransplant_j = 0 if crop is just after transplant and TS_i + TxDev_j < TSTR + StrsChoc_i
StopTransplant_j = 1 otherwise

TxDev_j = TxDev_j * StopTransplant_j
```

The internal cumulative development state is:

```text
TS_i(j) = TS_i(j-1) + TxDev_j
```

Stage thresholds are stored in `CTstade_i,s`. The crop advances when:

```text
TS_i(j) >= DVS_i
```

and then:

```text
Currstge_i(j) = Currstge_i(j-1) + 1
DVS_i <- DVS_i + CTstade_i,Currstge
```

The normalized development index used by LAI routines is:

```text
DVSt_i(j) = DVSt_i(j-1)
          + [TDV_i,s - TDV_i,s-1] * TxDev_j / CTstade_i,s
```

### 6.3 Oryza-like normalized phenology

In `pheno_sigmaT`, the code uses a normalized development rate:

```text
TS_i(j) = TS_i(j-1) + HU_j
DVS_i(j) = DVS_i(j-1) + DVR_i(j-1) * HU_j
```

Then `DVR` is stage-specific:

```text
stage 1: DVR = CTstade_i,1
stage 2: DVR = CTstade_i,2 * PPFAC_j
stage 3: DVR = CTstade_i,3
stage 4: DVR = CTstade_i,4
stage 5: DVR = CTstade_i,5
```

The crop stage is the interval containing `DVS_i`.

### 6.4 Mortality from drought and cold

The crop may die before maturity due to:

1. prolonged wilting signal
2. prolonged cold

The wilting signal is counted at soil level and compared to `NJFletri_i`.

Cold death is triggered when consecutive days below `Tcold_i` reach `NDieCold_i`.

## 7. Nitrogen stress on crop demand

Two nitrogen stress modes exist.

### 7.1 Seasonal mode

In `stressAzoteOld`:

```text
QNut_i = StockN + Kmo * ApportMO + ApportMin + Nsymb_i
NRF_i(j) = min(QNut_i / IFertMax_i, 1)
```

with `Kmo = 0.25`.

### 7.2 Daily mode

In `stressAzote`:

```text
NavailCult_i(j) = Navail_j + Nsymbjour_i
```

When biomass is already positive:

```text
NCvE_i = NCvEmin_i + (NCvEmax_i - NCvEmin_i) * alphaN_i
NUPTtarget_i(j) = Biom_i(j-1) / NCvE_i
NRF_i(j) = min(NavailCult_i(j) / NUPTtarget_i(j), 1)
```

Otherwise:

```text
NRF_i(j) = 1
```

The code therefore uses a ratio of available N to target uptake as the daily stress factor.

## 8. Water and nitrogen stress factors for growth

Two stress coefficients are used.

### 8.1 LAI stress factor

For leaf expansion:

```text
TurfacH_j = 1                                       if ContrainteW > 1 - SeuilTurg_i or no water stress
TurfacH_j = ContrainteW / (1 - SeuilTurg_i)         otherwise

TurfacN_j = 1                                       if NRF_i(j) = 1 or no N stress
TurfacN_j = NRF_i(j)                                otherwise

Turfac_j = min(TurfacH_j, TurfacN_j)
```

### 8.2 Biomass stress factor

For biomass production:

```text
WSfactH_j = 1                                       if ContrainteW > 1 - SeuilWS_i or no water stress
WSfactH_j = ContrainteW / (1 - SeuilWS_i)           otherwise

WSfactN_j = 1                                       if NRF_i(j) = 1 or no N stress
WSfactN_j = NRF_i(j)                                otherwise

WSfact_j = min(WSfactH_j, WSfactN_j)
```

## 9. LAI dynamics

### 9.1 Density correction

The density effect function is:

```text
deltaidens_i(j) = densite_j
```

If previous LAI exceeds the competition threshold `Laicomp_i` and density exceeds `bdens_i`:

```text
deltaidens_i(j) = densite_j * (densite_j / bdens_i) ^ adens_i
```

### 9.2 Light competition between crops

Competition is activated when:

```text
Nbcult > 1
LaiMC(j-1) >= 0.1
both crops are established
```

The dominated crop gets:

```text
CompFac_i(j) = exp(-k_other * LAI_other(j-1))
```

The dominant crop keeps:

```text
CompFac_i(j) = 1
```

### 9.3 Pre-senescence LAI growth

Before stage 3, CELSIUS uses an empirical logistic-like formulation.

Auxiliary variable:

```text
Ulai = 1 + (Vlaimax - 1) * DVSt / 0.4          for stage 1
Ulai = Vlaimax + (3 - Vlaimax) * (DVSt - 0.4) / 0.25   for stage 2
```

Potential increment:

```text
dLAIpot_i(j) = DLAImax_i / [1 + exp(5.5 * (Vlaimax - Ulai))] * HU_j * deltaidens_i(j)
```

Actual increment:

```text
dLAI_i(j) = dLAIpot_i(j) * Turfac_j * CompFac_i(j)
```

Then:

```text
LAI_i(j) = LAI_i(j-1) + dLAI_i(j)
```

In transplanting transitions, a density rescaling term is applied:

```text
LAI_i(j) = dLAI_i(j) + LAI_i(j-1) * dens_j / dens_j-1
```

### 9.4 Senescence

In the semi-arid variant currently used by the controller, after vegetative expansion:

```text
dLAISen_base = (LAIrec_i - LAI_i(JourSen)) / (2 - DVSt_i(JourSen))
```

Stress-accelerated senescence:

```text
dLAISen_i(j) = dLAISen_base - SensiSen_i * (1 - Turfac_j) * LAI_i(j-1)
dLAISen_i(j) = dLAISen_i(j) * [DVSt_i(j) - DVSt_i(j-1)]
```

Then:

```text
LAI_i(j) = LAI_i(j-1) + dLAISen_i(j)
LAI_i(j) = max(LAI_i(j), 0)
```

The implemented formulation therefore allows strong post-flowering stress to sharply accelerate canopy decline.

## 10. Radiation interception and biomass

### 10.1 CO2 effect

V32 initializes atmospheric CO₂ from `ListPAnnexes.CO2c`, then looks up the
simulation start year in `CO2Yearly` only when `Codcc = "0"`. When that year is
found, its `CO2` value replaces the station value; otherwise the model retains
the `ListPAnnexes` fallback and writes a warning. When `Codcc <> "0"`, the
yearly lookup is not performed.

`CO2Yearly` is indexed only by `yearCO2`, not by `idDclim`. The selected annual
concentration is consequently identical for every site with the same start
year. It is selected once during initialization and is not updated annually
inside a multi-year simulation.

The species-specific crop response uses `PlantSpecies.alphaCO2`:

```text
FCO2_i = 2 - exp( ln(2 - alphaCO2_i) * (CO2c - 350) / (600 - 350) )
```

This gives the two calibration points:

```text
FCO2_i(350 ppm) = 1
FCO2_i(600 ppm) = alphaCO2_i
```

For illustration, at 400 ppm the factor is approximately 1.044 for
`alphaCO2 = 1.2` (typical C3 value in the code comments) and 1.021 for
`alphaCO2 = 1.1` (typical C4 value).

### 10.2 Intercepted radiation

The code uses:

```text
raint_i(j) = 0.95 * ParSurRg * Rg_j * [1 - exp(-k_i * LAI_i(j))]
```

On transplant transition day, LAI may be scaled by the density ratio:

```text
raint_i(j) = 0.95 * ParSurRg * Rg_j * [1 - exp(-k_i * LAI_i(j) * dens_j / dens_j-1)]
```

### 10.3 Temperature response of radiation conversion

The temperature factor is a quadratic response around `tcopt_i`:

```text
Ftemp_i(j) = 1 - ((Tmoy_j - tcopt_i) / (tcmin_i - tcopt_i))^2     if Tmoy_j <= tcopt_i
Ftemp_i(j) = 1 - ((Tmoy_j - tcopt_i) / (tcmax_i - tcopt_i))^2     if Tmoy_j > tcopt_i
Ftemp_i(j) = max(Ftemp_i(j), 0)
```

### 10.4 Daily biomass increment

The crop biomass increment is:

```text
dBiom_i(j) = FCO2_i
           * WSfact_i(j)
           * PlantPReducFact
           * [Ebmax_i * raint_i(j) - 0.0815 * raint_i(j)^2]
           * Ftemp_i(j) / 100
```

Then:

```text
Biom_i(j) = Biom_i(j-1) + dBiom_i(j)
```

This is a radiation-use efficiency formulation with:

- intercepted radiation
- a quadratic penalty at high `raint`
- stress multiplier
- thermal multiplier
- CO2 multiplier

There is no other direct CO₂ term in the current implementation. CO₂ does not
directly modify LAI, phenology, transpiration, stomatal conductance or
water-use efficiency. It affects N uptake, grain number and yield indirectly
because these processes use the CO₂-modified biomass increment or accumulated
biomass.

### 10.5 Nitrogen uptake implied by biomass increment

The code also infers uptake from growth:

```text
NCvE_i = NCvEmin_i + (NCvEmax_i - NCvEmin_i) * alphaN_i
Nuptake_i(j) = dBiom_i(j) / NCvE_i
SigmaNuptake_i(j) = SigmaNuptake_i(j-1) + Nuptake_i(j)
```

## 11. Grain yield formation

### 11.1 Grain number

At phenological stage 4 onset:

```text
Vitmoy_i = 100 / Nbjgrain_i * sum_{n = JourDrp - Nbjgrain + 1}^{JourDrp} dBiom_i(n)
```

Then grain number:

```text
Ngrains_i = int(Cgrain_i * Vitmoy_i + Cgrainv0_i)
Ngrains_i = max(Ngrains_i, 0)
```

A per-plant upper bound is then applied:

```text
if Ngrains_i / dens_i(j) > Ngrmax_i:
    Ngrains_i = Ngrmax_i * dens_i(j)
```

### 11.2 Harvest index and grain mass

During grain filling:

```text
IR_i(j) = min(Vitircarb_i * (j - JourDrp_i + 1), IRmax_i)
```

Grain yield:

```text
Grain_i(j) = min(Biom_i(j) * IR_i(j), P1grainMax_i * Ngrains_i / 100)
```

Single grain weight:

```text
P1grain_i(j) = 100 * Grain_i(j) / Ngrains_i      if Ngrains_i > 0
P1grain_i(j) = -999.9                             otherwise
```

## 12. Root growth

Root front advance is:

```text
deltaZrac_i(j) = min(ZoneHumSousRacines_j, DeltaRacMax_i * HU_j)
```

with a correction for non-dominant root depth in mixtures:

```text
if Zrac_i(j-1) < ZracMC(j-1):
    ZoneHumSousRacines_j <- ZoneHumSousRacines_j + ZracMC(j-1) - Zrac_i(j-1)
```

Then:

```text
Zrac_i(j) = min(Zrac_i(j-1) + deltaZrac_i(j), Zracmax_i)
```

## 13. Mixed canopy processes

### 13.1 Mixed LAI and root depth

The canopy class aggregates individual crops:

```text
LAIMC_j = sum_i LAI_i(j)
ZRacineMC_j = max_i Zrac_i(j)
```

### 13.2 Potential evaporation below the canopy

The code uses:

```text
delta = k_extinction_MC - 0.2
EoSM_j = ETp_j * exp(-delta * LAIMC_j)
```

### 13.3 Crop coefficient and mixed transpiration

When canopy exists:

```text
Kc_j = 1 + (KmaxMC - 1) / [1 + exp(-1.5 * (LAIMC_j - 3))]
eo_j = ETp_j * Kc_j
eop_j = (eo_j - EoSM_j) * [beta + (1 - beta) * (Esol_j + Emulch_j) / EoSM_j]
```

with `beta = 1.4`.

Potential transpiration:

```text
TPotMC_j = eop_j
```

Water stress on canopy transpiration:

```text
WStressTMC_j = 1                         if ContrainteW > 0.7
WStressTMC_j = ContrainteW / 0.7         otherwise
```

because `PfactorMC = 0.7`.

Actual mixed transpiration:

```text
TranspiMC_j = TPotMC_j * WStressTMC_j
```

## 14. Mulch decomposition, interception, and runoff

### 14.1 Mulch mass and cover

Daily mulch mass:

```text
Qpaillis_j = Qpaillis_j-1 * exp(-Alpha_pail)
```

If mulch is applied that day:

```text
Qpaillis_j <- Qpaillis_j + QpaillisApport
```

Cover fraction:

```text
FracSoilCover_j = 1 - exp(-Beta_pail * Qpaillis_j)
```

### 14.2 Potential mulch evaporation

```text
Eomulch_j = EoSM_j * [1 - exp(-gamma_mulch * Qpaillis_j)]
```

### 14.3 Mulch water balance

The mulch store `Smulch` is first rescaled by biomass decay:

```text
Smulch_j^- = Smulch_j-1 * Qpaillis_j / Qpaillis_j-1
```

Evaporation has two terms.

Evaporation due to decomposition shrinkage:

```text
epail1_j = Smulch_j^- * (Qpaillis_j-1 - Qpaillis_j) / Qpaillis_j
epail1_j = max(epail1_j, 0)
```

Complementary evaporation:

```text
epail2_j = min(Eomulch_j - epail1_j, Smulch_j^-)
```

Total mulch evaporation:

```text
Emulch_j = epail1_j + epail2_j
```

Interception:

```text
intercep_j = precip_j * FracSoilCover_j
Smulch_j = Smulch_j^- - epail2_j + intercep_j
```

If storage exceeds capacity:

```text
if Smulch_j > CapaciteWMulch * Qpaillis_j:
    EauVersSol_j = precip_j - intercep_j + Smulch_j - CapaciteWMulch * Qpaillis_j
    Smulch_j = CapaciteWMulch * Qpaillis_j
else:
    EauVersSol_j = precip_j - intercep_j
```

### 14.4 Runoff

The runoff threshold is:

```text
seuil_j = seuil_ruis                                  if Albergel parameters are zero
seuil_j = (Ap4 - Ap2 * IKJ_j) / (Ap1 + Ap3 * IKJ_j)   otherwise
```

Then:

```text
Ruis_j = (Ap1 + Ap3 * IKJ_j + b_ruis * Qpaillis_j) * (precip_j - seuil_j)
Ruis_j = max(Ruis_j, 0)
```

If the surface type activates canopy protection:

```text
Ruis_j <- Ruis_j * exp(-0.5 * LAI_j)
```

Antecedent rainfall index update:

```text
IKJ_j+1 = (IKJ_j + precip_j) * exp(-0.5)
```

Water remaining after runoff:

```text
Eau_vers_Mulch_j = precip_j - Ruis_j
```

## 15. Soil hydraulic properties

The code computes average available water capacity from layer data.

For each layer `c`:

```text
TAW_total += (hcc_c - hmin_c) * da_c * epc_c / 10
```

Then:

```text
TAW = TAW_total / Ztotsol
```

The surface evaporation reservoir capacity is:

```text
TEW = (hcc_1 - hmin_1) * da_1 * Zsurf / 10
```

This is effectively a bucket model with:

- `Stsurf`
- `Strac`
- `Stnonrac`
- `Stprofond`

## 16. Soil evaporation

Potential soil evaporation below mulch:

```text
Eos_j = EoSM_j - Eomulch_j
```

Actual soil evaporation:

```text
Esol_j = Eos_j * ContrainteWSurf_j / SeuilEvap      if ContrainteWSurf_j < SeuilEvap
Esol_j = Eos_j                                      otherwise

Esol_j = min(Stsurf_j-1, Esol_j)
```

So surface water stress acts linearly below the threshold `SeuilEvap`.

## 17. Soil water balance

### 17.1 Surface store

Before evaporation and transpiration removal:

```text
bil = Stsurf_j-1 + precip_j
Stsurf_j = min(bil, TEW)
```

Then:

```text
Stsurf_j <- Stsurf_j - Esol_j
Stsurf_j <- Stsurf_j - min(Zsurf / Zrac_j, 1) * Transpi_j    if Zrac_j > 0
Stsurf_j = max(Stsurf_j, 0)
```

Surface water stress:

```text
ContrainteWSurf_j = Stsurf_j / TEW
```

### 17.2 Germination layer stock

When roots are absent:

```text
Stger_j = min(Stger_j-1 + precip_j, TAW * Zger) - Esol_j * Zger / Zsurf   if Zger < Zsurf
Stger_j = min(Stger_j-1 + precip_j, TAW * Zger) - Esol_j                  otherwise
Stger_j = max(Stger_j, 0)
```

### 17.3 Rooted zone stock

The rooted reservoir receives precipitation and newly colonized water:

```text
bil = precip_j + Strac_j-1 + Deltazrac_j * TAW
```

Then:

```text
Strac_j = min(bil, Zrac_j * TAW)
precip_j^* = max(bil - Zrac_j * TAW, 0)
```

Evaporation part extracted from rooted zone:

```text
E_Srac_j = Esol_j * min(1, Zrac_j / Zsurf)
Strac_j <- Strac_j - E_Srac_j - Transpi_j
Strac_j = max(Strac_j, 0)
```

### 17.4 Non-rooted but rootable store

```text
bil = Stnonrac_j-1 + precip_j^* - Deltazrac_j * TAW
```

If `Zrac < Zsurf`, part of evaporation is also taken from the non-rooted compartment:

```text
bil <- bil - Esol_j + E_Srac_j
```

Then:

```text
if bil > TAW * (Zracmax - Zrac_j):
    Dr_j = bil - TAW * (Zracmax - Zrac_j)
    Stnonrac_j = TAW * (Zracmax - Zrac_j)
else:
    Dr_j = 0
    Stnonrac_j = max(0, bil)
```

### 17.5 Deep store and drainage below maximum rootable depth

```text
Stprofond_j = min(Stprofond_j-1 + Dr_j, TAW * (Ztotsol - Zracmax))
Drprofmax_j = max(Stprofond_j-1 + Dr_j - TAW * (Ztotsol - Zracmax), 0)
```

### 17.6 Total stock and water stress

```text
Stocksol_j = Stprofond_j + Stnonrac_j + Strac_j
ContrainteW_j = Strac_j / (TAW * Zrac_j)      if Zrac_j > 0
ContrainteW_j = 1                              otherwise
```

### 17.7 Wilting signal

The soil-level wilting counter is:

```text
if Zrac_j > 0 and Stnonrac_j <= 1e-4 and Strac_j <= 0.02:
    SignalFletrissement_j = SignalFletrissement_j-1 + 1     if previous day also dry
else:
    SignalFletrissement_j = 0
```

This counter is then used by phenology and death logic.

## 18. Nitrogen balance in the soil

The daily N module only runs when daily N stress mode is active.

### 18.1 Soil temperature proxy

The model sets:

```text
Tsoil_j = Tmoy_j
```

This is a strong simplification.

### 18.2 Mineralization of soil organic N

From `MinNorgSS`:

```text
Tsoil_j = min(Tsoil_j, 35)
KnMin_j = 24 * exp(17.753 - 6350.5 / (Tsoil_j + 273))
Tfact_j = 1 - exp(-KnMin_j / 168)
```

Moisture factor:

```text
Wfact_j = 1.111 * ContrainteWSurf_j         if ContrainteWSurf_j < 0.9
Wfact_j = 10 - 10 * ContrainteWSurf_j       otherwise
Wfact_j = max(Wfact_j, 0)
```

Mineralized N:

```text
Nmin_j = Nhum_j-1 * Tfact_j * Wfact_j
Nmin_j = Nmin_j * (0.0002 - ConcNsol_j) / 0.0002 * 10
Nmin_j = max(Nmin_j, 0)
Nhum_j = Nhum_j-1 - Nmin_j
```

This means mineralization is shut down when soil solution N concentration approaches `0.0002`.

### 18.3 Denitrification

From `DenitNavail`:

```text
XConcNsol_j = min(ConcNsol_j, 0.0004)    if ContrainteWSurf_j = 1
KDenit_j = 6 * exp(0.07735 * Tsoil_j - 6.593)
NDenit_j = XConcNsol_j * (1 - exp(-KDenit_j))
NDenit_j = NDenit_j * Stsurf_j * 1000 * 10
```

### 18.4 Mineralization of applied organic matter

The routine `MinNorgapporteStics` is flagged in code comments as unfinished. Still, the implemented equations are:

```text
FTR_j = 12 / (1 + 51.6 * exp(-0.103 * Tsoil_j))
FH_j = 0.2 + 0.8 * ContrainteWSurf_j
```

Residue decomposition parameters:

```text
Kres = Akres + Bkres / CsurNRes
Hres = 1 - AHres + CsurNRes / (BHres + CsurNRes)
CNBio = AWB + BWB / CsurNRes
```

Main carbon and nitrogen flows:

```text
DCres_j = -Kres * Cres_j * FTR_j * FH_j * FN
DNres_j = DCres_j / CsurNRes

DCbio_j = -Yres * DCres_j - Kbio * Cbio_j-1 * FTR_j * FH_j
DNbio_j = -(Yres * DCres_j / (CNBio * Fbio)) - Kbio * Nbio_j-1 * FTR_j * FH_j

DChum_j = Kbio * Hres * FTR_j * FH_j
DNhum_j = DChum_j / CsurNhum

DNMinOrg_j = -DNres_j - DNhum_j - DNbio_j
```

This branch should be treated cautiously because the code itself says it is not stabilized.

### 18.5 Available mineral N pool

From `StockNavail`:

```text
Navail_j = Navail_j-1
         + Nmin_j
         + AppMinN_j
         - 0.35 * AppMinN_j
         + DNMinOrg_j
         - perteNDrain_j
         - NuptMC_j
         - 0.01 * Navail_j-1
         - NDenit_j
```

Then:

```text
Navail_j = max(Navail_j, 0)
```

This equation includes:

- soil organic N mineralization
- mineral fertilizer input
- a fixed 35% fertilizer loss term
- organic matter mineralization
- drainage loss
- plant uptake
- a fixed 1% gaseous loss term
- denitrification

### 18.6 Soil solution N concentration

The code uses:

```text
VolEausol_j = ZracmaxMC_j * TAW
ConcNsol_j = Navail_j / VolEausol_j
```

This is a uniform dilution across the whole rootable water volume.

### 18.7 N loss by drainage

```text
perteNDrain_j = ConcNsol_j-1 * Dr_j
```

## 19. Reference evapotranspiration helper functions

The helper module contains Penman-Monteith-like and Penman 1948 functions, but in the current CELSIUS run they are not the main internal ET engine. The daily ET forcing is read from the climate table.

Still, the code includes:

- astronomical radiation `Ra`
- FAO-like ET0 function `ET0pm`
- Penman-like `ET0pen48`

These may matter if the workflow later recomputes ET instead of reading it from `Dweather`.

## 20. Water stress pathways

Water stress in CELSIUS is not a single isolated module. It is propagated from the soil water balance toward emergence, leaf expansion, biomass accumulation, canopy transpiration, and plant death logic.

### 20.1 Soil control variables

The soil water balance produces two main control variables:

```text
ContrainteWSurf_j = Stsurf_j / TEW

ContrainteW_j = Strac_j / (TAW * Zrac_j)   if Zrac_j > 0
ContrainteW_j = 1                          otherwise
```

These are computed in `SolClass.EauSol` and represent:

- surface-layer dryness seen by soil evaporation and part of the N module
- root-zone dryness seen by crops and mixed-canopy transpiration

### 20.2 Emergence gating

Before roots are established, the model builds a germination-layer storage:

```text
Stger_j = min(Stger_j-1 + precip_j, TAW * Zger) - Esol_j * Zger / Zsurf    if Zger < Zsurf
Stger_j = min(Stger_j-1 + precip_j, TAW * Zger) - Esol_j                   otherwise
```

Then a Boolean emergence water constraint is triggered with hard thresholds:

```text
ContrainteHlevee_j = False   if Stger_j >= 0.14 * TAW * Zger
ContrainteHlevee_j = True    if Stger_j <= 0.10 * TAW * Zger
```

`PLanteClass.GerminLevee` allows simulated emergence only when:

```text
not ContrainteHlevee_j
```

So the first hydric effect of the model is not a continuous stress factor but an on/off gate on emergence.

### 20.3 Leaf expansion pathway

For canopy development, root-zone water status is converted into a leaf-expansion reduction factor:

```text
TurfacH_i(j) = 1                                   if ContrainteW_j > 1 - SeuilTurg_i
TurfacH_i(j) = ContrainteW_j / (1 - SeuilTurg_i)   otherwise
```

Then:

```text
Turfac_i(j) = min(TurfacH_i(j), TurfacN_i(j))
```

`Turfac_i(j)` multiplies daily LAI expansion and also increases senescence intensity in the semi-arid routine through:

```text
dLAISen_i(j) <- dLAISen_i(j) - SensiSen_i * (1 - Turfac_i(j)) * LAI_i(j-1)
```

So water stress affects both:

- reduced new leaf area
- accelerated leaf loss after the maximum-LAI phase

### 20.4 Biomass pathway

Biomass growth uses a separate hydric reduction factor:

```text
WSfactH_i(j) = 1                                  if ContrainteW_j > 1 - SeuilWS_i
WSfactH_i(j) = ContrainteW_j / (1 - SeuilWS_i)    otherwise
WSfact_i(j) = min(WSfactH_i(j), WSfactN_i(j))
```

Then biomass increment becomes:

```text
dBiom_i(j) = FCO2_i
           * WSfact_i(j)
           * [Ebmax_i * raint_i(j) - 0.0815 * raint_i(j)^2]
           * Ftemp_i(j) / 100
```

This means water deficit affects growth after radiation interception has been computed, by reducing radiation conversion rather than directly reducing intercepted radiation.

### 20.5 Mixed-canopy transpiration pathway

At the canopy scale, transpiration is reduced by another threshold function:

```text
WStressTMC_j = 1                     if ContrainteW_j > PfactorMC
WStressTMC_j = ContrainteW_j / PfactorMC   otherwise

TranspiMC_j = TPotMC_j * WStressTMC_j
```

With the current implementation:

```text
PfactorMC = 0.7
```

So canopy transpiration begins to decline earlier than leaf expansion or biomass in parameterizations where `1 - SeuilTurg_i` or `1 - SeuilWS_i` differ from `0.7`.

### 20.6 Severe drought and death signal

The soil routine also builds a severe-drought counter:

```text
if Zrac_j > 0 and Stnonrac_j <= 0.0001 and Strac_j <= 0.02:
    SignalFletrissement_j = SignalFletrissement_j-1 + 1
else:
    SignalFletrissement_j = 0
```

This signal is then passed to phenology:

```text
if SignalFletrissement_j >= NJFletri_i and crop still before early stage threshold:
    Die_i = True
    cropsta_i(j) = 0
```

So hydric stress can stop the crop in two ways:

- gradual slowdown through `TurfacH`, `WSfactH`, and `WStressTMC`
- abrupt failure through `SignalFletrissement`

### 20.7 Overall pathway summary

```text
Rain / irrigation
  -> runoff + mulch partitioning
  -> Stsurf, Strac, Stnonrac, Stprofond
  -> ContrainteWSurf and ContrainteW
  -> emergence gate + LAI reduction + biomass reduction + transpiration reduction
  -> possible wilting signal accumulation
  -> next-day canopy size, transpiration, and soil water
```

## 21. Main feedback loops

The model behavior is best understood through five feedback loops.

### 21.1 Water stress loop

```text
Rain / irrigation
  -> runoff + mulch partitioning
  -> soil stores
  -> ContrainteW / ContrainteWSurf
  -> Turfac / WSfact
  -> LAI, biomass, transpiration
  -> next day soil water
```

### 21.2 Phenology-climate loop

```text
Temperature + day length
  -> heat units + photoperiod factor
  -> stage progression
  -> LAI and yield phase changes
```

### 21.3 Canopy-radiation loop

```text
LAI
  -> intercepted radiation
  -> biomass growth
  -> further LAI growth
```

### 21.4 Nitrogen loop

```text
Soil N + fertilizers + organic matter
  -> Navail
  -> NRF
  -> LAI and biomass stress
  -> uptake
  -> updated Navail
```

### 21.5 Mixture competition loop

```text
Crop 1 LAI vs crop 2 LAI
  -> CompFac
  -> differential LAI growth
  -> altered canopy dominance
```

## 22. Main implementation caveats

Several equations above come directly from code comments or legacy branches and should be treated carefully.

The main caveats are:

- `MinNorgapporteStics` is explicitly marked unfinished.
- `ListResidus` is a required V32 input for the applied-residue branch and is
  included in the V32 export/build pipeline; this does not remove the need to
  validate the unfinished mineralization formulation.
- Soil layering is only partly represented; the actual water balance behaves mostly like a lumped bucket model.
- Soil temperature is approximated by air temperature.
- The crop mixture competition routine is acknowledged in comments as needing further checking.
- Some recursive multi-year logic is present but only partially operational.

## 23. Practical reading order

For process-level understanding, read in this order:

1. [SimulationControlClass.vb](../Converted/SimulationControlClass.vb)
2. [PLanteClass.vb](../Converted/PLanteClass.vb)
3. [CultureClass.vb](../Converted/CultureClass.vb)
4. [MulchClass.vb](../Converted/MulchClass.vb)
5. [SolClass.vb](../Converted/SolClass.vb)
6. [DataClimClass.vb](../Converted/DataClimClass.vb)

This is the closest thing to an executable scientific specification of the current model.
