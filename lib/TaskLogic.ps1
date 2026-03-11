
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
        [string]$worksheetName
    )

    $today = Get-Date -Format 'yyyy-MM-dd'
    $todayDate = [DateTime]::Parse($today).Date
    $dueTasks     = [System.Collections.Generic.List[hashtable]]::new()
    $upcomingTasks = [System.Collections.Generic.List[hashtable]]::new()
    switch ($worksheetName) {
        "Weekly Tasks" {$dueSoon = $todayDate.AddDays(1)}
        "Annual Tasks" {$dueSoon = $todayDate.AddDays(28)}
        "6-Monthly Tasks" {$dueSoon = $todayDate.AddDays(28)}
        default {$dueSoon = $todayDate.AddDays(7)}
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

#endregion
