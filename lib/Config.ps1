
#region Configuration Functions

function Load-Config {
    $config = @{}
    if (-not (Test-Path $configFile)) { return $config }

    try {
        $lines = Get-Content $configFile -ErrorAction Stop
    } catch {
        Write-Log "Could not read config file: $_" "WARN"
        return $config
    }

    if (-not $lines) { return $config }

    foreach ($line in $lines) {
        if ($line -match '^([^=]+)=(.*)$') {
            $config[$matches[1].Trim()] = $matches[2].Trim()
        }
    }
    return $config
}

function Save-Config {
    param(
        [string]$FilePath,
        [string]$Worksheet,
        [string]$DateColumn,
        [string]$DescriptionColumn,
        [string[]]$WorksheetNames
    )

    $config = Load-Config
    $upWorksheet = Get-WsEnvName $Worksheet

    $config["SPREADSHEET_PATH"] = $FilePath
    $config["WORKSHEET_NAMES"] = $WorksheetNames -join ','
    $config["${upWorksheet}_DATE_COLUMN"] = $DateColumn
    $config["${upWorksheet}_DESCRIPTION_COLUMN"] = $DescriptionColumn

    $configLines = $config.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }
    $configLines | Set-Content $configFile

    Write-Log "$Worksheet configuration saved to $configFile"
}

#endregion
