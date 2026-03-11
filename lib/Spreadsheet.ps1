
#region File and Spreadsheet Functions

function Select-SpreadsheetFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select your spreadsheet"
    $fileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv|All files (*.*)|*.*"

    if ($fileDialog.ShowDialog() -eq 'OK') {
        return $fileDialog.FileName
    }
    return $null
}

function Load-Spreadsheet {
    param(
        [string]$FilePath,
        [string]$WorksheetName = $null
    )

    # CSV: no COM objects involved, handle separately
    if ($FilePath -match '\.csv$') {
        try {
            $data    = Import-Csv $FilePath -ErrorAction Stop
            $headers = $data[0].PSObject.Properties.Name
            $rows    = @(,$headers) + @($data | ForEach-Object {
                $row = [System.Collections.Generic.List[object]]::new($headers.Count)
                foreach ($header in $headers) { $row.Add($_.$header) }
                ,$row.ToArray()
            })
            return @{ Data = $rows; WorksheetNames = @() }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not load file: $_", "Error", 'OK', 'Error')
            return @{ Data = $null; WorksheetNames = @() }
        }
    }

    # Excel: COM objects must be released even if an exception occurs
    $excel    = $null
    $workbook = $null
    $result   = @{ Data = $null; WorksheetNames = @() }
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false

        $workbook   = $excel.Workbooks.Open($FilePath, 0, $true)
        $wsNameList = [System.Collections.Generic.List[string]]::new()
        for ($i = 1; $i -le $workbook.Worksheets.Count; $i++) {
            $wsNameList.Add($workbook.Worksheets.Item($i).Name)
        }
        $worksheetNames = $wsNameList.ToArray()

        if ($WorksheetName -and $worksheetNames -contains $WorksheetName) {
            $result = @{
                Data           = Get-WorksheetData -Sheet $workbook.Worksheets.Item($WorksheetName)
                WorksheetNames = $worksheetNames
            }
        } elseif ($worksheetNames.Count -eq 1) {
            $result = @{
                Data           = Get-WorksheetData -Sheet $workbook.Worksheets.Item(1)
                WorksheetNames = $worksheetNames
            }
        } else {
            $result = @{ Data = $null; WorksheetNames = $worksheetNames }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not load file: $_", "Error", 'OK', 'Error')
        $result = @{ Data = $null; WorksheetNames = @() }
    } finally {
        # Guaranteed cleanup — runs on both success and exception paths
        if ($workbook) { try { $workbook.Close($false) } catch {} }
        if ($excel)    { try { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    }

    return $result
}

function Get-WorksheetData {
    param($Sheet)

    $usedRange = $Sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    $data = [System.Collections.Generic.List[object]]::new()
    for ($row = 1; $row -le $rowCount; $row++) {
        $rowData = [System.Collections.Generic.List[object]]::new($colCount)
        $hasData = $false
        for ($col = 1; $col -le $colCount; $col++) {
            $cellValue = $usedRange.Cells.Item($row, $col).Text
            if ($cellValue) { $hasData = $true }
            $rowData.Add($cellValue)
        }
        if ($hasData) {
            $data.Add($rowData.ToArray())
        }
    }

    if ($data.Count -gt 0) {
        $headerList = [System.Collections.Generic.List[string]]::new($data[0].Count)
        for ($i = 0; $i -lt $data[0].Count; $i++) {
            if ([string]::IsNullOrWhiteSpace($data[0][$i])) {
                $headerList.Add("Column_$($i + 1)")
            } else {
                $headerList.Add($data[0][$i].ToString().Trim())
            }
        }
        $data[0] = $headerList.ToArray()
    }

    return $data
}

function New-ExampleSpreadsheet {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title    = "Create example spreadsheet"
    $saveDialog.Filter   = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "My Tasks"
    if ($saveDialog.ShowDialog() -ne 'OK') { return $null }

    $path  = $saveDialog.FileName
    $today = Get-Date

    if ($path -match '\.xlsx$') {
        $excel = $null
        try {
            $sheetTasks = @{
                "Weekly Tasks"    = @(@{ Desc = "Submit weekly status update"; Days = 7 }, @{ Desc = "Team meeting preparation"; Days = 7 })
                "Monthly Tasks"   = @(@{ Desc = "Update project documentation"; Days = 30 })
                "Quarterly Tasks" = @(@{ Desc = "Review quarterly report"; Days = 90 })
                "6-Monthly Tasks" = @(@{ Desc = "Complete risk assessment"; Days = 180 })
                "Annual Tasks"    = @(@{ Desc = "Annual compliance training"; Days = 365 })
            }
            $names = @("Weekly Tasks", "Monthly Tasks", "Quarterly Tasks", "6-Monthly Tasks", "Annual Tasks")
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false; $excel.DisplayAlerts = $false
            $wb = $excel.Workbooks.Add()
            # Trim workbook down to 1 default sheet
            while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

            $isFirst = $true
            foreach ($name in $names) {
                if ($isFirst) {
                    $ws = $wb.Worksheets.Item(1)
                    $isFirst = $false
                } else {
                    $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
                }
                $ws.Name = $name

                # Header row
                $ws.Cells.Item(1, 1).Value2 = "Task Description"
                $ws.Cells.Item(1, 2).Value2 = "Date Last Performed"
                $ws.Cells.Item(1, 3).Value2 = "Due Date"
                $ws.Rows.Item(1).Font.Bold  = $true

                $tasks = $sheetTasks[$name]
                for ($i = 0; $i -lt $tasks.Count; $i++) {
                    $row = $i + 2
                    $rMax = -[int]($tasks[$i].Days * 0.9)
                    $rMin = -($tasks[$i].Days + 5)
                    $daysAgo = Get-Random -Maximum $rMax -Minimum $rMin
                    $ws.Cells.Item($row, 1).Value2  = $tasks[$i].Desc
                    $ws.Cells.Item($row, 2).Value2  = $today.AddDays($daysAgo).ToString("dd/MM/yyyy")
                    $ws.Cells.Item($row, 3).Formula = "=IF(ISBLANK(B$row), """", TEXT(IF(ISNUMBER(B$row), B$row, DATEVALUE(B$row))+$($tasks[$i].Days), ""dd/mm/yyyy""))"
                }
                $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
            }

            $wb.SaveAs($path, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
            $wb.Close($false)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not create Excel file: $_`n`nCreating CSV instead.", "Warning", 'OK', 'Warning') | Out-Null
            $path = [System.IO.Path]::ChangeExtension($path, '.csv')
            $csvLines = @("Task Description,Date Last Performed,Due Date")
            foreach ($t in $tasks) {
                $csvLines += "$($t.Desc),$($today.ToString('dd/MM/yyyy')),$($today.AddDays($t.Days).ToString('dd/MM/yyyy'))"
            }
            $csvLines | Set-Content $path
        } finally {
            if ($excel) { try { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
        }
    } else {
        $csvLines = @("Task Description,Date Last Performed,Due Date")
        foreach ($t in $tasks) {
            $csvLines += "$($t.Desc),$($today.ToString('dd/MM/yyyy')),$($today.AddDays($t.Days).ToString('dd/MM/yyyy'))"
        }
        $csvLines | Set-Content $path
    }

    return $path
}

#endregion
