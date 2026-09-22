param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [string]$EntryPoint = "Zyva",

    [switch]$Visible,

    [string[]]$ClearOutputTables = @("OutputSynt", "OutputD_1", "OutputD_2"),

    [switch]$KeepCompletionMessage
)

$ErrorActionPreference = "Stop"

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $Visible.IsPresent
    # Applies only to this automation instance; it does not change the user's
    # persistent Trust Center configuration.
    $access.AutomationSecurity = 1
    $access.OpenCurrentDatabase($AccessPath, $false)

    foreach ($table in $ClearOutputTables) {
        [void]$access.CurrentProject.Connection.Execute("DELETE FROM [$table]")
    }

    if (-not $KeepCompletionMessage) {
        $temporaryModule = Join-Path ([System.IO.Path]::GetTempPath()) (([guid]::NewGuid()).ToString() + ".bas")
        $access.SaveAsText(5, "Principal", $temporaryModule)
        $moduleText = Get-Content -LiteralPath $temporaryModule -Raw -Encoding Default
        $patchedText = $moduleText -replace '(?im)^\s*MsgBox\s*\(MsgFin\)\s*$', "' Completion message suppressed by batch runner"
        if ($patchedText -eq $moduleText) {
            throw "Could not locate the Principal completion MsgBox."
        }
        Set-Content -LiteralPath $temporaryModule -Value $patchedText -Encoding Default
        $access.DoCmd.SetWarnings($false)
        $access.DoCmd.DeleteObject(5, "Principal")
        $access.LoadFromText(5, "Principal", $temporaryModule)
        $access.DoCmd.SetWarnings($true)
    }

    # Principal updates this form's caption while running. Opening it hidden
    # keeps the batch run non-interactive while preserving the Access context.
    foreach ($form in $access.CurrentProject.AllForms) {
        if ($form.Name -eq "MenuPrincipal") {
            $access.DoCmd.OpenForm("MenuPrincipal", 0, "", "", 0, 1)
            break
        }
    }

    try {
        $access.DoCmd.SetWarnings($false)
    }
    catch {
    }

    $result = $access.Run($EntryPoint)
    Write-Output ("RUN_OK`t{0}`t{1}" -f $EntryPoint, $result)
}
catch {
    Write-Output ("ERROR`t{0}" -f $_.Exception.Message)
    exit 1
}
finally {
    if ($temporaryModule -and (Test-Path -LiteralPath $temporaryModule)) {
        Remove-Item -LiteralPath $temporaryModule -Force
    }
    if ($null -ne $access) {
        try {
            $access.DoCmd.SetWarnings($true)
        }
        catch {
        }

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
