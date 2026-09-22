param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string[]]$Tables = @(
        "Cultivars",
        "CO2Yearly",
        "Dweather",
        "FertiMin_List",
        "FertiOrga_List",
        "General_Parameters",
        "Irrigation_List",
        "ListPAnnexes",
        "ListResidus",
        "Mulch",
        "OptionsModel",
        "ParamIni",
        "PlantSpecies",
        "RuissellementObs",
        "SimUnitList",
        "Soil",
        "Soil_layers",
        "StadePheno",
        "Tech_Commun",
        "Tech_perCrop",
        "TypeSurfSol"
    )
)

$ErrorActionPreference = "Stop"

function Escape-TabValue {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [single] -or $Value -is [double]) {
        $text = $Value.ToString("R", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    elseif ($Value -is [decimal]) {
        $text = $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    elseif ($Value -is [datetime]) {
        $text = $Value.ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    else {
        $text = [string]$Value
    }
    $text = $text.Replace('"', '""')

    if ($text.Contains("`t") -or $text.Contains("`r") -or $text.Contains("`n")) {
        return '"' + $text + '"'
    }

    return $text
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$connection = New-Object -ComObject ADODB.Connection

try {
    $connection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;Persist Security Info=False;")

    foreach ($table in $Tables) {
        $target = Join-Path $OutputDir ($table + ".tsv")
        $recordset = New-Object -ComObject ADODB.Recordset
        $recordset.Open("SELECT * FROM [$table]", $connection)

        $writer = New-Object System.IO.StreamWriter($target, $false, [System.Text.Encoding]::UTF8)
        try {
            $header = @()
            for ($i = 0; $i -lt $recordset.Fields.Count; $i++) {
                $header += Escape-TabValue $recordset.Fields.Item($i).Name
            }
            $writer.WriteLine(($header -join "`t"))

            while (-not $recordset.EOF) {
                $values = @()
                for ($i = 0; $i -lt $recordset.Fields.Count; $i++) {
                    $values += Escape-TabValue $recordset.Fields.Item($i).Value
                }
                $writer.WriteLine(($values -join "`t"))
                $recordset.MoveNext()
            }
        }
        finally {
            $writer.Close()
            $recordset.Close()
        }

        Write-Output ("EXPORTED`t{0}" -f $table)
    }
}
finally {
    if ($connection.State -eq 1) {
        $connection.Close()
    }
}
