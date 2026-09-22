# CelsiusCli

VB .NET command-line port of the CELSIUS Access/VBA crop model.

## Build

```bash
DOTNET_CLI_HOME=/tmp dotnet build CelsiusCli/CelsiusCli.vbproj
```

## Run

```bash
DOTNET_CLI_HOME=/tmp dotnet run --project CelsiusCli/CelsiusCli.vbproj --no-build -- /path/to/model.sqlite
```

## Help

```bash
dotnet run --project CelsiusCli/CelsiusCli.vbproj -- --help
```

## Windows Command Install

To install a Windows command such as `celsius`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\celsius_pipeline\install_windows_command.ps1 `
  -ProjectDir D:\path\to\celsiusCli `
  -AddToUserPath
```

After opening a new terminal:

```powershell
celsius --help
celsius D:\path\to\celsius_model_input.db
```

## Windows Note

The current project uses different SQLite bindings on Linux and Windows.
If you switch to Windows after building on Linux or WSL, delete `bin` and `obj`
first so the Linux `SQLite.Interop.dll` is not reused:

```powershell
Remove-Item -Recurse -Force D:\path\to\celsiusCli\bin, D:\path\to\celsiusCli\obj
dotnet restore D:\path\to\celsiusCli\CelsiusCli.vbproj
dotnet run --project D:\path\to\celsiusCli\CelsiusCli.vbproj -- D:\path\to\celsius_model_input.db
```

The SQLite database must contain the model input tables with the same names used in the Access database, including:

- `SimUnitList`
- `General_Parameters`
- `OptionsModel`
- `ListPAnnexes`
- `Dweather`
- `ParamIni`
- `Tech_Commun`
- `Tech_perCrop`
- `Irrigation_List`
- `FertiMin_List`
- `FertiOrga_List`
- `Cultivars`
- `PlantSpecies`
- `StadePheno`
- `Soil`
- `Soil_layers`
- `Mulch`
- `RuissellementObs`
- `TypeSurfSol`

The CLI creates or reuses these output tables:

- `OutputSynt`
- `OutputD_1`
- `OutputD_2`

## Access to SQLite

1. Export the Access tables to TSV:

```bash
powershell.exe -NoProfile -File D:\docs\amei_workshop\sensitivity_analysis_outputs\export_access_tables.ps1 -AccessPath D:\Mes Donnees\TCMP\ACME_EspaceDeTravail\ACME_EspaceDeTravail\celsius\CelsiusV3nov17_dataArise.accdb -OutputDir D:\docs\amei_workshop\sensitivity_analysis_outputs\celsius_tsv_export
```

2. Build the SQLite file:

```bash
python3 build_sqlite_from_csv.py celsius_table_schema.csv celsius_tsv_export celsius_model_input.sqlite
```
