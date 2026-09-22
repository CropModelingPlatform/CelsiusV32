param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [Parameter(Mandatory = $true)]
    [string]$TargetProjectDir,

    [string]$WorkDir,

    [string]$TemplateProjectDir,

    [string]$SchemaPath,

    [string]$PythonExe = "python",

    [string]$DotnetExe = "dotnet",

    [switch]$BuildProject
)

$ErrorActionPreference = "Stop"

# Windows PowerShell 5.1 can evaluate default parameter expressions before
# $PSScriptRoot is initialized. Resolve path-dependent defaults here instead.
if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path $PSScriptRoot "work"
}
if ([string]::IsNullOrWhiteSpace($TemplateProjectDir)) {
    $TemplateProjectDir = Join-Path (Split-Path -Parent $PSScriptRoot) "CelsiusCli"
}
if ([string]::IsNullOrWhiteSpace($SchemaPath)) {
    $SchemaPath = Join-Path $WorkDir "celsius_table_schema.csv"
}

$vbaExportDir = Join-Path $WorkDir "vba_export"
$inputTsvDir = Join-Path $WorkDir "input_tsv"
$sqlitePath = Join-Path $WorkDir "celsius_model_input.db"

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Initialize-TargetProject {
    param(
        [string]$TemplateDir,
        [string]$TargetDir
    )

    if (Test-Path -LiteralPath $TargetDir) {
        Write-Output ("TARGET_EXISTS`t{0}" -f $TargetDir)
        return
    }

    $targetParent = Split-Path -Parent $TargetDir
    if ($targetParent) {
        Ensure-Directory -Path $targetParent
    }

    Copy-Item -LiteralPath $TemplateDir -Destination $TargetDir -Recurse

    $binDir = Join-Path $TargetDir "bin"
    $objDir = Join-Path $TargetDir "obj"
    if (Test-Path -LiteralPath $binDir) {
        Remove-Item -LiteralPath $binDir -Recurse -Force
    }
    if (Test-Path -LiteralPath $objDir) {
        Remove-Item -LiteralPath $objDir -Recurse -Force
    }

    Write-Output ("TARGET_CREATED`t{0}" -f $TargetDir)
}

Ensure-Directory -Path $WorkDir

Initialize-TargetProject -TemplateDir $TemplateProjectDir -TargetDir $TargetProjectDir

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "export_access_objects.ps1") `
    -AccessPath $AccessPath `
    -ExportDir $vbaExportDir

& $PythonExe (Join-Path $PSScriptRoot "convert_vba_to_vbnet.py") `
    $vbaExportDir `
    $TargetProjectDir

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "export_access_tables.ps1") `
    -AccessPath $AccessPath `
    -OutputDir $inputTsvDir

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "export_access_schema.ps1") `
    -AccessPath $AccessPath `
    -OutputPath $SchemaPath

& $PythonExe (Join-Path $PSScriptRoot "build_sqlite_from_tsv.py") `
    $SchemaPath `
    $inputTsvDir `
    $sqlitePath

if ($BuildProject) {
    & $DotnetExe build (Join-Path $TargetProjectDir "CelsiusCli.vbproj")
}

Write-Output ("SYNC_OK`t{0}" -f $sqlitePath)
