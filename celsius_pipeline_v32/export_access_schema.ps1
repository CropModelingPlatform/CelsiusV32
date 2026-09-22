param(
    [Parameter(Mandatory = $true)]
    [string]$AccessPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AccessPath, $false)
    $database = $access.CurrentDb()
    $result = @()

    $daoToAdoType = @{
        1 = 11   # Boolean
        2 = 2    # Byte
        3 = 2    # Integer
        4 = 3    # Long
        6 = 4    # Single
        7 = 5    # Double
        8 = 7    # DateTime
        10 = 130 # Text
        12 = 203 # Memo
        15 = 72  # GUID
        16 = 3   # BigInt
        17 = 17  # VarBinary
        18 = 130 # Char
        19 = 130 # Numeric
        20 = 5   # Decimal
        101 = 203 # Attachment/complex field fallback
    }

    foreach ($table in $database.TableDefs) {
        $tableName = [string]$table.Name
        if (-not $tableName.StartsWith("MSys") -and -not $tableName.StartsWith("~")) {
            for ($ordinal = 0; $ordinal -lt $table.Fields.Count; $ordinal++) {
                $field = $table.Fields.Item($ordinal)
                $dataType = $daoToAdoType[[int]$field.Type]
                if ($null -eq $dataType) {
                    $dataType = 130
                }
                $result += [PSCustomObject]@{
                    TableName  = $tableName
                    ColumnName = [string]$field.Name
                    Ordinal    = $ordinal + 1
                    DataType   = $dataType
                    IsNullable = -not [bool]$field.Required
                }
            }
        }
    }

    $result |
        Sort-Object TableName, Ordinal |
        Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Output ("SCHEMA_OK`t{0}`t{1}" -f $result.Count, $OutputPath)
}
finally {
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
    }
}
