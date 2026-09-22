param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectDir,

    [string]$InstallDir = "$env:LOCALAPPDATA\CelsiusCli",

    [string]$CommandName = "celsius",

    [string]$Configuration = "Release",

    [string]$Runtime = "win-x64",

    [switch]$SelfContained,

    [switch]$AddToUserPath
)

$ErrorActionPreference = "Stop"

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Add-DirectoryToUserPath {
    param([string]$PathToAdd)

    $current = [Environment]::GetEnvironmentVariable("Path", "User")
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($current)) {
        $parts = $current.Split(";") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    if ($parts -contains $PathToAdd) {
        Write-Output ("PATH_EXISTS`t{0}" -f $PathToAdd)
        return
    }

    $newPath = (($parts + $PathToAdd) | Select-Object -Unique) -join ";"
    [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    Write-Output ("PATH_ADDED`t{0}" -f $PathToAdd)
}

$publishDir = Join-Path $InstallDir "app"
$commandPath = Join-Path $InstallDir ($CommandName + ".cmd")
$projectFile = Join-Path $ProjectDir "CelsiusCli.vbproj"

Ensure-Directory -Path $InstallDir
if (Test-Path -LiteralPath $publishDir) {
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}
Ensure-Directory -Path $publishDir

$publishArgs = @(
    "publish",
    $projectFile,
    "-c", $Configuration,
    "-r", $Runtime,
    "--output", $publishDir
)

if ($SelfContained) {
    $publishArgs += "--self-contained"
    $publishArgs += "true"
}
else {
    $publishArgs += "--self-contained"
    $publishArgs += "false"
}

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw ("dotnet publish a echoue avec le code de sortie {0}. La commande n'a pas ete installee." -f $LASTEXITCODE)
}

$publishedExe = Join-Path $publishDir "CelsiusCli.exe"
if (-not (Test-Path -LiteralPath $publishedExe)) {
    throw ("Publication incomplete : executable introuvable ({0})." -f $publishedExe)
}

$cmdContent = @"
@echo off
"$publishDir\CelsiusCli.exe" %*
"@

Set-Content -LiteralPath $commandPath -Value $cmdContent -Encoding ASCII
Write-Output ("COMMAND_CREATED`t{0}" -f $commandPath)

if ($AddToUserPath) {
    Add-DirectoryToUserPath -PathToAdd $InstallDir
}

Write-Output ("INSTALL_OK`t{0}" -f $commandPath)
