
#region Date Parsing and Task Checking

function Parse-Date {
    param($DateValue)

    if ([string]::IsNullOrWhiteSpace($DateValue)) {
        return $null
    }

    $dateFormats = @(
        'yyyy-MM-dd HH:mm:ss',
        'yyyy-MM-dd',
        'dd/MM/yyyy',
        'dd-MM-yyyy',
        'dd MMMM yyyy',
        'dd MMM yyyy',
        'd/M/yyyy',
        'yyyy/MM/dd',
        'MMMM dd, yyyy',
        'MMM dd, yyyy',
        'MM/dd/yyyy',
        'MM-dd-yyyy',
        'M/d/yyyy'
    )

    $ukCulture = [System.Globalization.CultureInfo]::new('en-GB')

    foreach ($format in $dateFormats) {
        try {
            $parsedDate = [DateTime]::ParseExact($DateValue.ToString().Trim(), $format, $ukCulture)
            return $parsedDate.Date
        } catch {
            continue
        }
    }

    try {
        $parsedDate = [DateTime]::Parse($DateValue.ToString().Trim(), $ukCulture)
        return $parsedDate.Date
    } catch {
        return $null
    }
}

function Get-DueTasks {
    param(
        [array]$Data,
        [string]$DateColumn,
        [string]$DescriptionColumn,
        [string]$worksheetName,
        [int]$LookAheadDays = -1
    )

    $today = Get-Date -Format 'yyyy-MM-dd'
    $todayDate = [DateTime]::Parse($today).Date
    $dueTasks     = [System.Collections.Generic.List[hashtable]]::new()
    $upcomingTasks = [System.Collections.Generic.List[hashtable]]::new()

    if ($LookAheadDays -ge 0) {
        $dueSoon = $todayDate.AddDays($LookAheadDays)
    } else {
        switch -Wildcard ($worksheetName) {
            "*Weekly*"    { $dueSoon = $todayDate.AddDays(1) }
            "*Annual*"    { $dueSoon = $todayDate.AddDays(28) }
            "*6-Monthly*" { $dueSoon = $todayDate.AddDays(28) }
            default       { $dueSoon = $todayDate.AddDays(7) }
        }
    }

    $headers = $Data[0]
    try {
        $dateColIndex = [array]::IndexOf($headers, $DateColumn)
        $descColIndex = [array]::IndexOf($headers, $DescriptionColumn)

        if ($dateColIndex -eq -1 -or $descColIndex -eq -1) {
            [System.Windows.Forms.MessageBox]::Show("Column not found in header", "Error", 'OK', 'Error')
            return @()
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not find columns: $_", "Error", 'OK', 'Error')
        return @()
    }

    for ($i = 1; $i -lt $Data.Count; $i++) {
        $row = $Data[$i]
        if (-not $row -or $row.Count -eq 0) { continue }

        $taskDate = Parse-Date $row[$dateColIndex]
        $description = if ($row[$descColIndex]) { $row[$descColIndex].ToString() } else { "Task in row $($i + 1)" }

        if ($taskDate -and $taskDate -le $todayDate) {
            $daysOverdue = ($todayDate - $taskDate).Days
            $dueTasks.Add(@{
                Description = $description
                DueDate = $taskDate
                DaysOverdue = $daysOverdue
                RowIndex = $i + 1
            })
        } elseif ($taskDate -and $taskDate -gt $todayDate -and $taskDate -le $dueSoon) {
            $upcomingTasks.Add(@{
                Description = $description
                DueDate = $taskDate
                DaysUntilDue = ($taskDate - $todayDate).Days
                WeekdayDue = $taskDate.DayOfWeek
                RowIndex = $i + 1
            })
        } elseif (-not $taskDate -and $row[$dateColIndex]) {
            $dueTasks.Add(@{
                Description = $description
                DueDate = "Invalid Date"
                DaysOverdue = -1
                RowIndex = $i + 1
            })
        } elseif (-not $taskDate -and -not $row[$dateColIndex]) {
            $dueTasks.Add(@{
                Description = $description
                DueDate = "No due date"
                DaysOverdue = -1
                RowIndex = $i + 1
            })
        }
    }

    return @{
        Due      = @($dueTasks)
        Upcoming = @($upcomingTasks)
    }
}


function Get-BankHolidays {
    param(
        [string]$Region = 'england-and-wales'
    )

    if ($Region -eq 'disabled' -or [string]::IsNullOrWhiteSpace($Region)) { return @() }

    $regionMap = @{
        'england-and-wales' = 'england-and-wales'
        'scotland'          = 'scotland'
        'northern-ireland'  = 'northern-ireland'
    }
    $apiRegion = if ($regionMap.ContainsKey($Region)) { $regionMap[$Region] } else { 'england-and-wales' }
    $configKey = 'BANK_HOLIDAYS_' + ($apiRegion.ToUpper() -replace '-', '_')
    $today     = (Get-Date).Date

    # Load stored dates from config and prune any that are in the past
    $cfg    = Load-Config
    $raw    = if ($cfg[$configKey]) { @($cfg[$configKey] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) } else { @() }
    $future = @($raw | Where-Object {
        $dt = [DateTime]::MinValue
        [DateTime]::TryParseExact($_, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$dt) -and $dt -ge $today
    })

    if ($future.Count -lt $raw.Count) {
        $cfg[$configKey] = $future -join ','
        ($cfg.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) | Set-Content $configFile
        Write-Log "Pruned $($raw.Count - $future.Count) past bank holiday date(s) from config."
    }

    if ($future.Count -gt 0) { return $future }

    # Config is empty — fetch from the API and save future dates
    try {
        $json    = Invoke-RestMethod 'https://www.gov.uk/bank-holidays.json' -TimeoutSec 10
        $fetched = @($json.$apiRegion.events | ForEach-Object { $_.date } | Where-Object {
            $dt = [DateTime]::MinValue
            [DateTime]::TryParseExact($_, 'yyyy-MM-dd', $null, [System.Globalization.DateTimeStyles]::None, [ref]$dt) -and $dt -ge $today
        })
        $cfg[$configKey] = $fetched -join ','
        ($cfg.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) | Set-Content $configFile
        Write-Log "Bank holidays fetched from gov.uk and saved to config ($($fetched.Count) future dates)."
        return $fetched
    } catch {
        Write-Log "Could not fetch bank holidays: $_" "WARN"
        return @()
    }
}

#endregion
