param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [Parameter(Mandatory = $true)]
    [string]$ProjectDir,

    [string]$WorkDir,

    [string]$EntryPoint = "Zyva",

    [string]$SchemaPath,

    [string]$PythonExe = "python",

    [string]$DotnetExe = "dotnet",

    [double]$FloatTolerance = 1e-6
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path $PSScriptRoot "regression_work"
}
if ([string]::IsNullOrWhiteSpace($SchemaPath)) {
    $SchemaPath = Join-Path $WorkDir "celsius_table_schema.csv"
}

$accessRunDir = Join-Path $WorkDir "access_run"
$accessInputDir = Join-Path $WorkDir "access_input_tsv"
$accessOutputDir = Join-Path $WorkDir "access_output_tsv"
$sqliteOutputDir = Join-Path $WorkDir "sqlite_output_tsv"
$sqliteDbPath = Join-Path $WorkDir "celsius_regression_input.db"
$accessCopyPath = Join-Path $accessRunDir "reference_run.accdb"

function Write-RegressionStep {
    param([int]$Percent, [string]$Message)
    Write-Output ("REGRESSION {0}% - {1}" -f $Percent, $Message)
}

function Assert-ExternalCommand {
    param([string]$Step)
    if ($LASTEXITCODE -ne 0) {
        throw ("Etape echouee ({0}), code de sortie {1}." -f $Step, $LASTEXITCODE)
    }
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

Ensure-Directory -Path $WorkDir
Ensure-Directory -Path $accessRunDir

foreach ($generatedPath in @($accessInputDir, $accessOutputDir, $sqliteOutputDir)) {
    if (Test-Path -LiteralPath $generatedPath) {
        Remove-Item -LiteralPath $generatedPath -Recurse -Force
    }
}
if (Test-Path -LiteralPath $sqliteDbPath) {
    Remove-Item -LiteralPath $sqliteDbPath -Force
}

Write-RegressionStep 0 "Preparation de la copie Access"
Copy-Item -LiteralPath $AccessPath -Destination $accessCopyPath -Force

Write-RegressionStep 5 "Execution du modele VBA Access (phase longue)"
& (Join-Path $PSScriptRoot "run_access_model.ps1") `
    -AccessPath $accessCopyPath `
    -EntryPoint $EntryPoint

Write-RegressionStep 40 "Export des tables Access"
& (Join-Path $PSScriptRoot "export_access_tables.ps1") `
    -AccessPath $accessCopyPath `
    -OutputDir $accessInputDir

& (Join-Path $PSScriptRoot "export_access_schema.ps1") `
    -AccessPath $accessCopyPath `
    -OutputPath $SchemaPath

$outputTables = @("OutputSynt", "OutputD_1", "OutputD_2")
& (Join-Path $PSScriptRoot "export_access_tables.ps1") `
    -AccessPath $accessCopyPath `
    -OutputDir $accessOutputDir `
    -Tables $outputTables

& $PythonExe (Join-Path $PSScriptRoot "build_sqlite_from_tsv.py") `
    $SchemaPath `
    $accessInputDir `
    $sqliteDbPath
Assert-ExternalCommand "construction SQLite"

Write-RegressionStep 55 "Compilation du CLI VB.NET"
& $DotnetExe build (Join-Path $ProjectDir "CelsiusCli.vbproj")
Assert-ExternalCommand "compilation VB.NET"

Write-RegressionStep 60 "Execution du modele VB.NET"
& $DotnetExe run --project (Join-Path $ProjectDir "CelsiusCli.vbproj") --no-build -- $sqliteDbPath
Assert-ExternalCommand "execution VB.NET"

Write-RegressionStep 92 "Export et comparaison des sorties"
& $PythonExe (Join-Path $PSScriptRoot "export_sqlite_tables.py") `
    $sqliteDbPath `
    $sqliteOutputDir
Assert-ExternalCommand "export des sorties SQLite"

& $PythonExe (Join-Path $PSScriptRoot "compare_tsv_tables.py") `
    $accessOutputDir `
    $sqliteOutputDir `
    --float-tolerance $FloatTolerance `
    --relative-tolerance 1e-6 `
    --ignore-column OutputSynt.BilanNnonOK `
    --ignore-column OutputD_1.DrainageON
Assert-ExternalCommand "comparaison des sorties"

Write-RegressionStep 100 "Test termine"
Write-Output "REGRESSION_OK"
