param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [Parameter(Mandatory = $true)]
    [string]$ExportDir
)

$ErrorActionPreference = "Stop"

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

try {
    Ensure-Directory -Path $ExportDir

    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AccessPath, $false)

    Write-Output "ACCESS_OK"

    foreach ($module in $access.CurrentProject.AllModules) {
        $target = Join-Path $ExportDir ($module.Name + ".bas")
        $access.SaveAsText(5, $module.Name, $target)
        Write-Output ("MODULE`t{0}" -f $module.Name)
    }

    foreach ($form in $access.CurrentProject.AllForms) {
        $target = Join-Path $ExportDir ($form.Name + ".form.txt")
        $access.SaveAsText(2, $form.Name, $target)
        Write-Output ("FORM`t{0}" -f $form.Name)
    }

    foreach ($report in $access.CurrentProject.AllReports) {
        $target = Join-Path $ExportDir ($report.Name + ".report.txt")
        $access.SaveAsText(3, $report.Name, $target)
        Write-Output ("REPORT`t{0}" -f $report.Name)
    }

    foreach ($query in $access.CurrentData.AllQueries) {
        $target = Join-Path $ExportDir ($query.Name + ".query.txt")
        $access.SaveAsText(1, $query.Name, $target)
        Write-Output ("QUERY`t{0}" -f $query.Name)
    }
}
catch {
    Write-Output ("ERROR`t{0}" -f $_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $access) {
        try {
            $access.CloseCurrentDatabase()
        }
        catch {
        }

        try {
            $access.Quit()
        }
        catch {
        }
    }
}
