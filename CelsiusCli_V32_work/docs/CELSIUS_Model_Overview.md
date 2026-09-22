# CELSIUS V32 Model Overview

This document explains CELSIUS V32 from the execution flow in the main program and the main simulation controller. It is meant to help someone read the code as a system, not just as isolated classes.

> **V32 revision status (September 2026).** This revision is tied to the
> sources in `CelsiusCli_V32_work` and to the Access/VBA export in
> `celsius_audit_new/access_objects`. The scientific descriptions remain
> implementation-oriented rather than an independent formal specification.

For a graphical view of the coupled processes, see the
[CELSIUS V32 biophysical process graph](CELSIUS_V32_Process_Graph.md).

The description below is based mainly on:

- [CelsiusRuntime.vb](../CelsiusRuntime.vb)
- [Principal.bas](../../celsius_audit_new/access_objects/Principal.bas)
- [SimulationControlClass.vb](../Converted/SimulationControlClass.vb)

## 1. What CELSIUS does

CELSIUS runs a list of simulation units. Each simulation unit defines:

- which climate series to use
- which soil to use
- which technical management to use
- which initial state to use
- which model options to use
- the start and end dates of the simulation

For each unit, the model:

1. loads all required inputs from database tables
2. initializes plant, soil, mulch, and management states
3. simulates day by day
4. writes summary outputs and daily outputs

The main simulation object is `SimulationControlClass`. It orchestrates the other domain classes. The scientific model is therefore not concentrated in one formula; it is a coordinated set of interacting process modules.

## 2. Top-Level Execution Flow

### 2.1 CLI / runtime entry

The .NET CLI entry point is [Program.vb](../Program.vb). It passes the SQLite database path to `CelsiusRuntime.Run`.

`CelsiusRuntime.Run`:

1. opens the SQLite database
2. ensures output tables exist
3. hands control to `PrincipalRunner.Run`

This is the VB .NET equivalent of the original Access VBA `Principal()` procedure in [Principal.bas](../../celsius_audit_new/access_objects/Principal.bas).

### 2.2 Principal runner

`PrincipalRunner.Run` implements the outer simulation loop:

1. read `SimUnitList`
2. clear `OutputSynt`
3. loop over each simulation row
4. for each row:
   - `ReadParameters`
   - `Simulation`
   - `SortieSynthesis`
   - `EcritDresu`
   - `MemoEtatFinal` if recursive mode is enabled

This means the conceptual structure is:

```text
Database
  -> SimUnitList
     -> one SimulationControlClass run per row
        -> load state and parameters
        -> simulate all days
        -> write outputs
```

## 3. Main Orchestrator: `SimulationControlClass`

`SimulationControlClass` is the hub of the model. It owns one instance of each main process class:

- `SimulationUnitClass`
- `DataClimClass`
- `GestionTechniqueClass`
- `PLanteClass`
- `CultureClass`
- `SolClass`
- `MulchClass`
- `OptionsModelClass`
- `GenParamClass`
- `EtatInitialClass`
- `EtatFinalClass`
- `IrrigClass`
- `FertiMinClass`
- `FertiOrgaClass`
- `ApportsOrgaClass`

Its two core methods are:

- `ReadParameters(...)`
- `Simulation()`

The model is easiest to understand if you read these two methods first.

## 4. Input Assembly Phase: `ReadParameters`

### 4.1 Simulation identity and dates

`SimulationUnitClass` reads one row from `SimUnitList` and turns it into the simulation identity:

- `IdSim`
- `IdTech_Com`
- `IdWeather`
- `IdSoil`
- `idIni`
- `idGenParam`
- `idCodModel`
- `StartYear`, `StartDay`, `EndYear`, `EndDay`
- computed `NbJourSimul`

This class defines the simulation window and the keys used to retrieve everything else.

### 4.2 General parameters and model options

`ReadParameters` then loads:

- `GenParamClass` from `General_Parameters`
- `OptionsModelClass` from `OptionsModel`

These are global switches and parameters, for example:

- water stress on/off
- nitrogen stress on/off
- emergence simulated or forced
- daily outputs on/off
- phenology mode
- altitude correction on/off
- automatic management on/off

Important consequence: many branches in the simulation depend on `OptionsModelClass`.

### 4.3 Climate forcing

`DataClimClass.LisClimD` builds the daily climate arrays from:

- `ListPAnnexes`
- `CO2Yearly`
- `Dweather`

It prepares arrays such as:

- `Tmin`, `Tmax`, `Tmoy`
- `Rg`
- `Etp`
- `Plu`
- `DOY`
- `DAP`
- `CurrentYear`
- `DAYL`

Climate is therefore preloaded for the whole simulation before the daily loop starts.

#### CO₂ selection and role

`DataClimClass.LisClimD` first reads the station-level fallback concentration
from `ListPAnnexes.CO2c`. The `Codcc` field of the current simulation unit then
controls whether the yearly table is consulted:

- when `Codcc = "0"`, the model searches `CO2Yearly` for
  `yearCO2 = StartYear` and uses the corresponding `CO2` value;
- when `Codcc <> "0"`, it keeps `ListPAnnexes.CO2c`;
- when `Codcc = "0"` but the start year is absent, it reports a warning and
  keeps `ListPAnnexes.CO2c`.

`CO2Yearly` has no `idDclim` field. Its annual value is therefore global: all
sites starting in the same year receive the same concentration. Only the
fallback in `ListPAnnexes` can vary by climate station. The lookup uses
`StartYear` once during initialization; the concentration is not updated during
a multi-year simulation.

For each crop, `PLanteClass.Iniplante` combines this concentration with the
species parameter `PlantSpecies.alphaCO2`:

```text
CO2fact_i = 2 - exp[ln(2 - alphaCO2_i) * (CO2c - 350) / 250]
```

Consequently, `CO2fact = 1` at 350 ppm and `CO2fact = alphaCO2` at 600 ppm.
The factor multiplies the daily biomass increment in `PLanteClass.biomasse`.
Its downstream effects on N uptake, grain number and yield pass through the
additional biomass. V32 does not apply a separate direct CO₂ effect to LAI,
phenology, transpiration, stomatal conductance or water-use efficiency.

### 4.4 Initial state

`EtatInitialClass` loads the initial conditions from `ParamIni`, unless recursive mode is active. In recursive mode it rebuilds the initial state from the previous `EtatFinal`.

This is the mechanism intended to chain simulations without fully resetting the system between them.

### 4.5 Technical management

`GestionTechniqueClass.LisTech` loads:

- `Tech_Commun`
- `Tech_perCrop`

This defines:

- number of crops in the association
- sowing date(s)
- transplanting settings
- mulch application
- irrigation on/off
- mineral fertilization on/off
- organic fertilization on/off
- automatic reseeding behavior
- observed runoff forcing on/off

This class is central because it turns static parameter tables into event calendars.

### 4.6 Irrigation and fertilization calendars

If the management switches are active:

- `IrrigClass` loads `Irrigation_List`
- `FertiMinClass` loads `FertiMin_List`
- `FertiOrgaClass` loads `FertiOrga_List`

If organic fertilization is active and `TypeMorga <> "indetermine"`, then `ApportsOrgaClass` is also loaded.

Important caveat:

- `ApportsOrgaClass` expects a `ListResidus` table
- `ListResidus` is part of the V32 input schema and is exported by
  `celsius_pipeline_v32`
- a V32 input database must therefore provide this table when the organic
  residue branch is enabled

### 4.7 Plant, soil, and mulch parameterization

Then `ReadParameters` assembles the core biophysical subsystems:

- `PLanteClass` from `Cultivars`, `PlantSpecies`, `StadePheno`
- `SolClass` from `Soil`, `Soil_layers`, and `TypeSurfSol`
- `MulchClass` from `Mulch`

It also optionally loads observed runoff from `RuissellementObs`.

### 4.8 Cross-initialization between modules

The final part of `ReadParameters` is important because it wires modules together:

1. `Plante.Iniplante(...)` initializes each crop using management dates, densities, soil limit, and CO2.
2. `Culture.InitMixedCanopy(...)` collects crop-specific properties into mixed-canopy properties.
3. `Sol.initsol(...)` initializes water storage from the initial state and root-zone geometry.
4. `Culture.lisCultureMC()` initializes mixed-canopy stress state.
5. `Mulch.LisMulch(...)` initializes mulch state.

So `ReadParameters` is not just reading tables. It builds a consistent simulation state across modules.

## 5. Daily Simulation Loop

The daily loop is in `SimulationControlClass.Simulation()`. This is the main process graph of CELSIUS.

At a high level, each day does:

```text
For each day:
  manage crops and technical events
  update living crops
  update mulch and runoff
  update soil evaporation and water balance
  update canopy transpiration
  update nitrogen balance if enabled
```

## 6. Detailed Process Order Within One Day

### 6.1 Crop management before growth

For each crop:

- seasonal N stress can be computed with `stressAzoteOld`
- automatic management can react via `CyberPlouck`
- automatic reseeding can shift sowing dates via `SemisAuto`
- a planned sowing day triggers `Plante.Iniplante` and activates the crop

This stage decides whether a crop exists and is active on the current day.

### 6.2 Crop-level development and growth

If at least one crop is alive, each living crop is updated.

Main sequence for each living crop:

1. carry previous crop state forward
2. if early stage: simulate germination/emergence with `GerminLevee`
3. if established:
   - apply daily nitrogen stress if enabled
   - update phenology with:
     - `pheno_sigmaT` for `Oryza`
     - `phenoCTphot` otherwise
   - update LAI with `Calcule_LAI_SemiAride`
   - update biomass with `biomasse`
   - update grain yield with `Rendement`
   - update root depth with `Croirac`
4. aggregate crop values into the mixed canopy:
   - `Culture.LaiMixedCanopy`
   - `Culture.CroiRacMixedCrop`
   - `Culture.NuptakeMixedCrop`

This is the main plant engine of the model.

### 6.3 Mulch and runoff

After crop updates:

1. `Mulch.BiomasseMulch`
   - decomposes existing mulch
   - optionally adds mulch on the mulch application day
2. `Culture.Evaporation_Pot_SolMulch`
   - computes the potential evaporation under canopy
3. `Mulch.Ruissellement`
   - computes runoff, unless observed runoff is forcing the value
4. `Mulch.BilanMulch`
   - partitions water between mulch storage and water reaching the soil

This means rainfall and irrigation go first through surface partitioning before the soil water balance.

### 6.4 Soil evaporation and canopy transpiration

Next:

1. `Sol.evaporation`
   - computes soil evaporation under current surface water stress
2. `Culture.CalcTranspiMC`
   - computes mixed-canopy transpiration from ETp, soil evaporation, mulch evaporation, and water stress
3. `Sol.EauSol`
   - updates the soil water stores using:
     - water from mulch
     - canopy transpiration
     - root-zone depth
     - root growth

This is the hydraulic coupling:

```text
Climate ETp + Rain
  -> canopy / mulch / runoff partitioning
  -> soil evaporation + transpiration demand
  -> soil water storage update
  -> water stress feedback to plants
```

### 6.5 Daily nitrogen balance

If nitrogen stress is active and `TypeNPKstress = 2`, the model also runs the daily N module:

1. `Sol.TempSoil`
2. `Sol.MinNorgSS`
3. `Sol.DenitNavail`
4. `Sol.MinNorgapporteStics`
5. `Sol.StockNavail`
6. `Sol.ConcNEauSol`

The daily N module therefore combines:

- mineralization from soil organic matter
- denitrification
- mineral fertilization
- organic fertilization
- crop N uptake
- N in rainfall
- drainage losses

This branch is the main nutrient feedback path into crop growth.

## 7. Interaction Map Between Main Classes

The most useful mental model is:

```text
SimulationUnit
  -> selects climate, soil, management, initial state, options

OptionsModel
  -> activates or deactivates branches

DataClim
  -> provides daily forcing: temperature, radiation, ETp, rain, daylength

GestionTechnique
  -> provides event calendars: sowing, transplanting, irrigation, fertilization, mulch

PLante
  -> simulates crop phenology, LAI, biomass, yield, roots, N demand and uptake

Culture
  -> aggregates individual crops into one mixed canopy

Mulch
  -> simulates surface cover, runoff, mulch evaporation, water transfer to soil

Sol
  -> simulates soil water storage, evaporation, drainage, and nitrogen dynamics
```

The most important interactions are:

- `DataClim -> PLante`
  - temperature, radiation, photoperiod, CO2
- `DataClim -> Mulch/Sol`
  - rain, ETp, temperature
- `GestionTechnique -> PLante`
  - sowing, transplanting, densities
- `GestionTechnique -> Irrig/Ferti/Mulch`
  - event activation and timing
- `PLante -> Culture`
  - LAI, roots, N uptake
- `Culture -> Sol`
  - transpiration demand, root depth
- `Mulch -> Sol`
  - water reaching the soil after runoff and interception
- `Sol -> PLante`
  - water stress, N availability, wilting signal

This is the core feedback structure of the model.

## 8. Outputs

The model writes two types of outputs.

### 8.1 Summary outputs

`SortieSynthesis` writes one row per simulation to `OutputSynt`, including:

- phenological dates
- death day
- biomass and grain yield
- max LAI
- cumulative soil evaporation
- cumulative drainage
- cumulative runoff
- cumulative mulch evaporation
- cumulative transpiration
- soil N stock

### 8.2 Daily outputs

`EcritDresu` writes daily outputs to:

- `OutputD_1`
- `OutputD_2`

depending on the crop index in the mixture.

These tables contain daily values for:

- climate
- crop development and biomass
- root depth
- soil water stores
- transpiration and evaporation
- runoff and mulch
- N concentration, drainage, N losses, uptake, and stress indicators

This is the most useful layer if you want to diagnose behavior process by process.

## 9. Main Behavioral Switches

Several flags strongly change the model path.

### 9.1 `OptionsModel`

Most important switches:

- `ActiveWstress`
- `ActiveNstress`
- `TypeNPKstress`
- `simlevee`
- `EcritDResus`
- `CorrigAlti`
- `CyberST`
- `CodeDevelop`

### 9.2 `GestionTechnique`

Most important switches:

- `IrrigON`
- `fertiminON`
- `fertiorgON`
- `DriveRuiObs`
- `RepiquageON`
- automatic reseeding fields

Together, these switches decide which submodels are active and which input tables matter for a given run.

## 10. How To Read the Code Efficiently

If you want to understand the model without getting lost, read in this order:

1. [CelsiusRuntime.vb](../CelsiusRuntime.vb)
2. [Principal.bas](../../celsius_audit_new/access_objects/Principal.bas)
3. [SimulationControlClass.vb](../Converted/SimulationControlClass.vb)
4. [SimulationUnitClass.vb](../Converted/SimulationUnitClass.vb)
5. [OptionsModelClass.vb](../Converted/OptionsModelClass.vb)
6. [GestionTechniqueClass.vb](../Converted/GestionTechniqueClass.vb)
7. [DataClimClass.vb](../Converted/DataClimClass.vb)
8. [PLanteClass.vb](../Converted/PLanteClass.vb)
9. [CultureClass.vb](../Converted/CultureClass.vb)
10. [MulchClass.vb](../Converted/MulchClass.vb)
11. [SolClass.vb](../Converted/SolClass.vb)

This order follows the actual control flow.

## 11. Current Caveats

- The current .NET port still keeps some VBA-era patterns and compatibility layers.
- Applied-residue mineralization in `MinNorgapporteStics` is still marked as
  unfinished in the inherited scientific code; V32 supplies its parameters
  through `ListResidus`, but the formulation still requires validation.
- V32 first reads the mean atmospheric CO₂ value from `ListPAnnexes`, then
  replaces it with the matching simulation-year value from `CO2Yearly` when
  available; the crop-level response is applied through the existing `FCO2`
  calculation.
- The model is written for up to two crops in association in many parts of the code.
- Several comments in the original code indicate unfinished or weakly tested branches, especially around recursive mode and some multi-layer soil logic.

## 12. Short Conceptual Summary

CELSIUS is a daily crop-soil-water-mulch model driven by database-defined simulation units.

Its architecture is:

```text
Simulation unit selection
  -> parameter loading
  -> system initialization
  -> daily coupled simulation
     -> crop development and growth
     -> canopy aggregation
     -> runoff and mulch water balance
     -> soil water balance
     -> optional soil-plant nitrogen balance
  -> summary and daily outputs
```

That is the main structure to keep in mind when extending or refactoring the code.
