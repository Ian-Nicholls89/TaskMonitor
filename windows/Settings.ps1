
function Show-SettingsWindow {
    param(
        [string]$FilePath,
        [string[]]$WorksheetNames,
        [hashtable]$SavedConfig,
        [System.Windows.Window]$OwnerWindow,
        [hashtable]$WsHeadersCache = @{}
    )

    # set registry variables to load or unload later
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    $isRegistered = [bool](Get-ItemProperty -Path $regPath -Name "TaskMonitor" -ErrorAction SilentlyContinue)

    # Use pre-loaded headers from cache (populated at startup and on every refresh)
    $wsHeaders = $WsHeadersCache

    # Build dynamic worksheet sections
    $worksheetSections = ""
    foreach ($ws in $WorksheetNames) {
        $wsKey  = "WS_" + ($ws -replace '[^a-zA-Z0-9]', '_')
        $wsEnv  = Get-WsEnvName $ws
        $headers = $wsHeaders[$ws]
        if (-not $headers) { continue }

        $escapedWsName = [System.Security.SecurityElement]::Escape($ws)
        $dateItems = Get-ComboItemsXaml -Columns $headers
        $descItems = $dateItems

        # Default look-ahead based on sheet name pattern (matches TaskLogic.ps1 defaults)
        $defaultLookAhead = switch -Wildcard ($ws) {
            "*Weekly*"    { 1 }
            "*Annual*"    { 28 }
            "*6-Monthly*" { 28 }
            default       { 7 }
        }
        $savedLookAhead = if ($SavedConfig["${wsEnv}_LOOKAHEAD_DAYS"]) {
            $SavedConfig["${wsEnv}_LOOKAHEAD_DAYS"]
        } else {
            $defaultLookAhead
        }

        $worksheetSections += @"
                        <TextBlock Text="$escapedWsName"
                                   FontSize="15" FontWeight="Bold" Foreground="#00BCD4" Margin="0,20,0,6"/>
                        <TextBlock Text="DUE BY column:" FontSize="13" Foreground="#B0B0B0" Margin="0,0,0,4"/>
                        <ComboBox x:Name="Date_$wsKey"
                                  Style="{StaticResource MaterialDesignOutlinedComboBox}"
                                  FontSize="13" Padding="8,10" Margin="0,0,0,12">
$dateItems
                        </ComboBox>
                        <TextBlock Text="TASK DESCRIPTION column:" FontSize="13" Foreground="#B0B0B0" Margin="0,0,0,4"/>
                        <ComboBox x:Name="Desc_$wsKey"
                                  Style="{StaticResource MaterialDesignOutlinedComboBox}"
                                  FontSize="13" Padding="8,10" Margin="0,0,0,12">
$descItems
                        </ComboBox>
                        <TextBlock Text="LOOK-AHEAD DAYS:" FontSize="13" Foreground="#B0B0B0" Margin="0,0,0,4"/>
                        <TextBox x:Name="LookAhead_$wsKey"
                                 Text="$savedLookAhead"
                                 Style="{StaticResource MaterialDesignOutlinedTextBox}"
                                 FontSize="13" Padding="8,10" Margin="0,0,0,4"
                                 Width="100" HorizontalAlignment="Left"/>
"@
    }

    $escapedFilePath = [System.Security.SecurityElement]::Escape($FilePath)

    [xml]$xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    Title="TaskMonitor - Settings"
    Icon="$icon"
    Width="520" SizeToContent="Height" MaxHeight="720"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    ShowTitleBar="False"
    UseNoneWindowStyle="True"
    Background="#1E1E2E"
    GlowBrush="#00BCD4"
    NonActiveGlowBrush="#333333">

    $(Get-WindowResourcesXaml)

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        $(Get-TitleBarXaml -Title 'Settings' -Buttons @(
            @{ Name='TitleCloseBtn'; Icon='&#xE8BB;'; Style='TitleBarCloseBtn' }
        ))

        <Grid Grid.Row="1" Margin="24,8,24,24">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Text="Settings"
                       FontSize="22" FontWeight="Light" Foreground="#E0E0E0"
                       Margin="0,0,0,16"/>

            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="520">
                <StackPanel>
                    <!-- File path -->
                    <TextBlock Text="SPREADSHEET FILE" FontSize="12" FontWeight="Bold"
                               Foreground="#888888" Margin="0,0,0,6"/>
                    <Grid Margin="0,0,0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="FilePathBox" Grid.Column="0"
                                 Text="$escapedFilePath"
                                 IsReadOnly="True"
                                 Style="{StaticResource MaterialDesignOutlinedTextBox}"
                                 FontSize="12" Foreground="#AAAAAA"
                                 Background="#2A2A3E" BorderBrush="#444444"
                                 VerticalContentAlignment="Center"
                                 Padding="8,10" Margin="0,0,8,0"/>
                        <Button x:Name="BrowseBtn" Grid.Column="1"
                                Content="Browse"
                                Style="{StaticResource MaterialDesignOutlinedButton}"
                                Foreground="#00BCD4" BorderBrush="#00BCD4"
                                FontSize="13" Height="46" Padding="12,0"
                                VerticalAlignment="Center"/>
                    </Grid>

                    <Separator Background="#333333" Margin="0,12,0,0"/>

                    <!-- Working days -->
                    <TextBlock Text="WORKING DAYS" FontSize="12" FontWeight="Bold"
                               Foreground="#888888" Margin="0,16,0,8"/>
                    <WrapPanel Orientation="Horizontal">
                        <CheckBox x:Name="ChkMonday"   Content="Mon" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkTuesday"  Content="Tue" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkWednesday" Content="Wed" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkThursday" Content="Thu" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkFriday"   Content="Fri" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkSaturday" Content="Sat" Foreground="#E0E0E0" Margin="0,0,16,6"/>
                        <CheckBox x:Name="ChkSunday"   Content="Sun" Foreground="#E0E0E0" Margin="0,0,0,6"/>
                    </WrapPanel>

                    <Separator Background="#333333" Margin="0,12,0,0"/>

                    <!-- Bank holidays -->
                    <TextBlock Text="BANK HOLIDAYS" FontSize="12" FontWeight="Bold"
                               Foreground="#888888" Margin="0,16,0,8"/>
                    <TextBlock Text="Highlight tasks due on public holidays in amber (like non-working days)."
                               FontSize="12" Foreground="#666688" TextWrapping="Wrap" Margin="0,0,0,8"/>
                    <ComboBox x:Name="BankHolidayRegion"
                              Style="{StaticResource MaterialDesignOutlinedComboBox}"
                              FontSize="13" Padding="8,10" Margin="0,0,0,4">
                        <ComboBoxItem Content="Disabled" Tag="disabled"/>
                        <ComboBoxItem Content="England &amp; Wales" Tag="england-and-wales"/>
                        <ComboBoxItem Content="Scotland" Tag="scotland"/>
                        <ComboBoxItem Content="Northern Ireland" Tag="northern-ireland"/>
                    </ComboBox>

                    <Separator Background="#333333" Margin="0,12,0,0"/>

                    <!-- Run at startup -->
                    <TextBlock Text="RUN AT STARTUP?" FontSize="12" FontWeight="Bold"
                               Foreground="#888888" Margin="0,16,0,8"/>
                    <WrapPanel Orientation="Horizontal">
                        <Controls:ToggleSwitch x:Name="StartupSwitch"
                            OffContent="No"
                            OnContent="Yes"/>
                    </WrapPanel>

                    <Separator Background="#333333" Margin="0,12,0,0"/>

                    <!-- Minimise / close to tray -->
                    <TextBlock Text="MINIMISE / CLOSE TO TRAY?" FontSize="12" FontWeight="Bold"
                               Foreground="#888888" Margin="0,16,0,8"/>
                    <WrapPanel Orientation="Horizontal">
                        <Controls:ToggleSwitch x:Name="TraySwitch"
                            OffContent="No"
                            OnContent="Yes"/>
                    </WrapPanel>

                    <Separator Background="#333333" Margin="0,12,0,0"/>

                    <!-- Worksheet column pickers -->
$worksheetSections

                </StackPanel>
            </ScrollViewer>

            <!-- Buttons -->
            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
                <Button x:Name="SaveBtn" Content="Save"
                        Style="{StaticResource MaterialDesignRaisedButton}"
                        Background="#00BCD4" BorderBrush="#00BCD4" Foreground="White"
                        FontSize="14" Width="130" Height="40"
                        materialDesign:ButtonAssist.CornerRadius="6"
                        Margin="0,0,16,0"/>
                <Button x:Name="CancelBtn" Content="Cancel"
                        Style="{StaticResource MaterialDesignOutlinedButton}"
                        Foreground="#EF5350" BorderBrush="#EF5350"
                        FontSize="14" Width="130" Height="40"
                        materialDesign:ButtonAssist.CornerRadius="6"/>
            </StackPanel>
        </Grid>
    </Grid>
</Controls:MetroWindow>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $settingsWin = [Windows.Markup.XamlReader]::Load($reader)
    $settingsWin.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)
    if ($OwnerWindow) { $settingsWin.Owner = $OwnerWindow }

    $titleBar   = $settingsWin.FindName("TitleBar")
    $titleBar.Add_MouseLeftButtonDown({ $settingsWin.DragMove() }.GetNewClosure())
    $titleCloseBtn = $settingsWin.FindName("TitleCloseBtn")
    $titleCloseBtn.Add_Click({ $settingsWin.Close() }.GetNewClosure())

    $filePathBox = $settingsWin.FindName("FilePathBox")
    $browseBtn   = $settingsWin.FindName("BrowseBtn")
    $saveBtn     = $settingsWin.FindName("SaveBtn")
    $cancelBtn   = $settingsWin.FindName("CancelBtn")
    $startupSwitch = $settingsWin.FindName("StartupSwitch")
    $startupSwitch.IsOn = $isRegistered
    $traySwitch = $settingsWin.FindName("TraySwitch")
    $traySwitch.IsOn = $SavedConfig['MINIMISE_TO_TRAY'] -eq 'True'

    $result = @{ Confirmed = $false; NewFilePath = $FilePath; WorksheetSettings = @{}; WorkingDays = $null; MinimiseToTray = $null; BankHolidayRegion = $null }

    # Pre-populate ComboBoxes
    foreach ($ws in $WorksheetNames) {
        $wsKey   = "WS_" + ($ws -replace '[^a-zA-Z0-9]', '_')
        $wsEnv   = Get-WsEnvName $ws
        $headers = $wsHeaders[$ws]
        if (-not $headers) { continue }
        $dateCombo = $settingsWin.FindName("Date_$wsKey")
        $descCombo = $settingsWin.FindName("Desc_$wsKey")
        $savedDate = $SavedConfig["${wsEnv}_DATE_COLUMN"]
        $savedDesc = $SavedConfig["${wsEnv}_DESCRIPTION_COLUMN"]
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -eq $savedDate) { $dateCombo.SelectedIndex = $i }
            if ($headers[$i] -eq $savedDesc) { $descCombo.SelectedIndex = $i }
        }
    }

    # Pre-populate working days checkboxes
    $dayMap = @{
        Monday    = 'ChkMonday';    Tuesday  = 'ChkTuesday';  Wednesday = 'ChkWednesday'
        Thursday  = 'ChkThursday'; Friday   = 'ChkFriday';   Saturday  = 'ChkSaturday'
        Sunday    = 'ChkSunday'
    }
    $savedWorkingDays = if ($SavedConfig['WORKING_DAYS']) {
        @($SavedConfig['WORKING_DAYS'] -split ',' | ForEach-Object { $_.Trim() })
    } else {
        @('Monday','Tuesday','Wednesday','Thursday','Friday')
    }
    foreach ($day in $dayMap.Keys) {
        $chk = $settingsWin.FindName($dayMap[$day])
        if ($chk) { $chk.IsChecked = $day -in $savedWorkingDays }
    }

    # Pre-populate bank holiday region
    $bankHolidayCombo = $settingsWin.FindName("BankHolidayRegion")
    $savedRegion = if ($SavedConfig['BANK_HOLIDAY_REGION']) { $SavedConfig['BANK_HOLIDAY_REGION'] } else { 'england-and-wales' }
    for ($bhi = 0; $bhi -lt $bankHolidayCombo.Items.Count; $bhi++) {
        if ($bankHolidayCombo.Items[$bhi].Tag -eq $savedRegion) {
            $bankHolidayCombo.SelectedIndex = $bhi
            break
        }
    }

    $browseBtn.Add_Click({
        $newPath = Select-SpreadsheetFile
        if ($newPath) {
            $filePathBox.Text = $newPath
            $result.NewFilePath = $newPath
        }
    }.GetNewClosure())

    $saveBtn.Add_Click({
        if ($startupSwitch.IsOn) {
            $cmd = "wscript.exe `"$directory\TaskMonitor.vbs`""
            Set-ItemProperty -Path $regPath -Name "TaskMonitor" -Value $cmd
        } else {
            Remove-ItemProperty -Path $regPath -Name "TaskMonitor" -ErrorAction SilentlyContinue
        }
        foreach ($ws in $WorksheetNames) {
            $wsKey   = "WS_" + ($ws -replace '[^a-zA-Z0-9]', '_')
            $headers = $wsHeaders[$ws]
            if (-not $headers) { continue }
            $dateCombo      = $settingsWin.FindName("Date_$wsKey")
            $descCombo      = $settingsWin.FindName("Desc_$wsKey")
            $lookAheadBox   = $settingsWin.FindName("LookAhead_$wsKey")
            $lookAheadVal   = -1
            if ($lookAheadBox -and [int]::TryParse($lookAheadBox.Text.Trim(), [ref]$lookAheadVal) -and $lookAheadVal -lt 0) {
                $lookAheadVal = 0
            }
            $result.WorksheetSettings[$ws] = @{
                DateCol      = if ($dateCombo.SelectedIndex -ge 0) { $headers[$dateCombo.SelectedIndex] } else { $null }
                DescCol      = if ($descCombo.SelectedIndex -ge 0) { $headers[$descCombo.SelectedIndex] } else { $null }
                LookAheadDays = $lookAheadVal
            }
        }
        $selectedRegionItem = $bankHolidayCombo.SelectedItem
        $result.BankHolidayRegion = if ($selectedRegionItem) { $selectedRegionItem.Tag } else { 'england-and-wales' }
        $checkedDays = @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday') | Where-Object {
            $chk = $settingsWin.FindName("Chk$_")
            $chk -and $chk.IsChecked
        }
        $result.WorkingDays = $checkedDays -join ','
        $result.MinimiseToTray = $traySwitch.IsOn
        $result.Confirmed = $true
        $settingsWin.Close()
    }.GetNewClosure())

    $cancelBtn.Add_Click({ $settingsWin.Close() }.GetNewClosure())

    $settingsWin.ShowDialog() | Out-Null
    return $result
}
