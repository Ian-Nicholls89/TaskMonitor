
#region WPF Display Functions

function New-TaskCountOverlay {
    param([int]$Count)

    $size = 32
    $bmp  = New-Object System.Drawing.Bitmap($size, $size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    # Red circle background
    $bg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(239, 83, 80))
    $g.FillEllipse($bg, 0, 0, $size - 1, $size - 1)

    # Count text
    $text     = if ($Count -gt 99) { '99+' } else { $Count.ToString() }
    $fontSize = if ($Count -gt 9)  { 13 }    else { 17 }
    $font     = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold)
    $sf       = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $fg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $g.DrawString($text, $font, $fg, [System.Drawing.RectangleF]::new(0, 0, $size, $size), $sf)

    $g.Dispose(); $font.Dispose(); $bg.Dispose(); $fg.Dispose()

    # Convert to WPF BitmapSource
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $ms.Position = 0
    $bi = New-Object System.Windows.Media.Imaging.BitmapImage
    $bi.BeginInit()
    $bi.StreamSource = $ms
    $bi.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bi.EndInit()
    $bi.Freeze()
    $ms.Dispose()

    return $bi
}

function Set-TaskComplete {
    param(
        [string]$FilePath,
        [string]$WorksheetName,
        [int]$RowIndex
    )
    try {
        $stream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
        $stream.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Please close the spreadsheet before marking a task as complete.`n`n$([System.IO.Path]::GetFileName($FilePath))",
            "File In Use", 'OK', 'Warning') | Out-Null
        return $false
    }

    $excel    = $null
    $workbook = $null
    $success  = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($FilePath, 0, $false)
        $sheet = $workbook.Worksheets.Item($WorksheetName)
        $sheet.Cells.Item($RowIndex, 2).Value2 = (Get-Date).ToString("dd MMM yyyy")
        $workbook.Save()
        $success = $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Could not update spreadsheet: $_", "Error", 'OK', 'Error') | Out-Null
    } finally {
        if ($workbook) { try { $workbook.Close($false) } catch {} }
        if ($excel)    { try { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {} }
    }
    return $success
}

function New-TaskItemsXaml {
    param(
        [array]$DueTasks,
        [array]$UpcomingTasks,
        [string[]]$WorkingDays = @('Monday','Tuesday','Wednesday','Thursday','Friday')
    )
    $taskItemsXaml = ""

    if ($DueTasks.Count -gt 0) {
        $taskItemsXaml += @"
                    <TextBlock Text="Due / Overdue Tasks ($($DueTasks.Count))"
                               FontSize="16" FontWeight="Bold" Foreground="#EF5350"
                               Margin="0,0,0,12"/>
"@
        for ($di = 0; $di -lt $DueTasks.Count; $di++) {
            $task = $DueTasks[$di]
            if ($task.DaysOverdue -eq 0) {
                $status = "DUE TODAY"
                $statusColor = "#FFA726"
                $borderColor = "#FFA726"
            } elseif ($task.DaysOverdue -eq -1) {
                $status = if ($task.DueDate -eq "Invalid Date") { "INVALID DATE" } else { "NO DUE DATE" }
                $statusColor = "#FFEE58"
                $borderColor = "#FFEE58"
            } else {
                $status = "OVERDUE by $($task.DaysOverdue) day(s)"
                $statusColor = "#EF5350"
                $borderColor = "#EF5350"
            }

            $dateDisplay = if ($task.DueDate -is [DateTime]) { $task.DueDate.ToString('dd/MM/yyyy') } else { $task.DueDate }
            $worksheetInfo = if ($task.Worksheet) { "$($task.Worksheet) - " } else { "" }
            $escapedDesc = [System.Security.SecurityElement]::Escape("$worksheetInfo$($task.Description)")
            $escapedDate = [System.Security.SecurityElement]::Escape($dateDisplay)
            $escapedStatus = [System.Security.SecurityElement]::Escape($status)

            switch ($task.DaysOverdue) {
                0       { $dueLabelXaml = "<Run Text=""DUE TODAY"" FontWeight=""Bold"" Foreground=""$statusColor""/>" }
                -1      { $dueLabelXaml = "<Run Text=""$escapedStatus"" FontWeight=""Bold"" Foreground=""$statusColor""/>" }
                default { $dueLabelXaml = "<Run Text=""OVERDUE by $($task.DaysOverdue) day(s)"" FontWeight=""Bold"" Foreground=""$statusColor""/><Run Text="" - due $escapedDate"" Foreground=""#888888""/>" }
            }

            $taskItemsXaml += @"
                    <Border BorderBrush="$borderColor" BorderThickness="0,0,0,2"
                            Background="#2A2A3E" CornerRadius="4" Padding="14,10" Margin="0,0,0,8">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0"
                                       Text="$escapedDesc"
                                       FontSize="14" Foreground="#E0E0E0" TextWrapping="Wrap"/>
                            <TextBlock Grid.Row="1" Grid.Column="0"
                                       FontSize="12" Margin="0,4,0,0">$dueLabelXaml</TextBlock>
                            <Button Grid.Row="0" Grid.Column="2" Grid.RowSpan="3"
                                Background="Transparent" BorderBrush="#69F0AE" BorderThickness="1"
                                Width="40" Height="40" Padding="0"
                                VerticalAlignment="Center" Margin="8,0,0,0"
                                Tag="D_$di">
                                <Button.Template>
                                    <ControlTemplate TargetType="Button">
                                        <Border x:Name="Bd"
                                                Background="{TemplateBinding Background}"
                                                BorderBrush="{TemplateBinding BorderBrush}"
                                                BorderThickness="{TemplateBinding BorderThickness}"
                                                CornerRadius="4">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="#2269F0AE"/>
                                                <Setter TargetName="Bd" Property="BorderBrush" Value="#A0EEC0"/>
                                            </Trigger>
                                            <Trigger Property="IsPressed" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="#4469F0AE"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Button.Template>
                                <TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="16" Foreground="#69F0AE"/>
                            </Button>
                        </Grid>
                    </Border>
"@
        }
    }

    if ($UpcomingTasks.Count -gt 0) {
        $taskItemsXaml += @"
                    <TextBlock Text="Upcoming Tasks ($($UpcomingTasks.Count))"
                               FontSize="16" FontWeight="Bold" Foreground="#00BCD4"
                               Margin="0,16,0,12"/>
"@
        for ($ui = 0; $ui -lt $UpcomingTasks.Count; $ui++) {
            $task = $UpcomingTasks[$ui]
            if ($task.DaysUntilDue -eq 1) {
                $status = "Due tomorrow"
            } else {
                $status = "Due in $($task.DaysUntilDue) days"
            }

            $isNonWorkingDay = ($null -ne $task.WeekdayDue) -and
                               ($WorkingDays.Count -gt 0) -and
                               ($task.WeekdayDue.ToString() -notin $WorkingDays)
            if ($isNonWorkingDay) {
                $upcomingBorderColor  = "#FFA726"
                $upcomingStatusColor  = "#FFA726"
                $weekdayLabel = [System.Security.SecurityElement]::Escape($task.WeekdayDue.ToString())
                $nonWorkingNote = "<Run Text="" ($weekdayLabel)"" Foreground=""$upcomingStatusColor""/>"
            } else {
                $upcomingBorderColor  = "#00BCD4"
                $upcomingStatusColor  = "#69F0AE"
                $nonWorkingNote = ""
            }

            $dateDisplay = if ($task.DueDate -is [DateTime]) { $task.DueDate.ToString('dd/MM/yyyy') } else { $task.DueDate }
            $worksheetInfo = if ($task.Worksheet) { "$($task.Worksheet) - " } else { "" }
            $escapedDesc   = [System.Security.SecurityElement]::Escape("$worksheetInfo$($task.Description)")
            $escapedDate   = [System.Security.SecurityElement]::Escape($dateDisplay)
            $escapedStatus = [System.Security.SecurityElement]::Escape($status)

            if ($task.DaysUntilDue -eq 1) {
                $dueLabelXaml = "<Run Text=""Due tomorrow"" FontWeight=""Bold"" Foreground=""$upcomingStatusColor""/>"
            } else {
                $dueLabelXaml = "<Run Text=""Due in $($task.DaysUntilDue) days"" FontWeight=""Bold"" Foreground=""$upcomingStatusColor""/><Run Text="" on $escapedDate"" Foreground=""#888888""/>$(if ($isNonWorkingDay) {$nonWorkingNote})"
            }

            $taskItemsXaml += @"
                    <Border BorderBrush="$upcomingBorderColor" BorderThickness="0,0,0,2"
                            Background="#2A2A3E" CornerRadius="4" Padding="14,10" Margin="0,0,0,8">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0"
                                       Text="$escapedDesc"
                                       FontSize="14" Foreground="#E0E0E0" TextWrapping="Wrap"/>
                            <TextBlock Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2"
                                       FontSize="12" Margin="0,4,0,0">$dueLabelXaml</TextBlock>
                            <Button Grid.Row="0" Grid.Column="2" Grid.RowSpan="2"
                                Background="Transparent" BorderBrush="#69F0AE" BorderThickness="1"
                                Width="40" Height="40" Padding="0"
                                VerticalAlignment="Center" Margin="8,0,0,0"
                                Tag="U_$ui">
                                <Button.Template>
                                    <ControlTemplate TargetType="Button">
                                        <Border x:Name="Bd"
                                                Background="{TemplateBinding Background}"
                                                BorderBrush="{TemplateBinding BorderBrush}"
                                                BorderThickness="{TemplateBinding BorderThickness}"
                                                CornerRadius="4">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="#2269F0AE"/>
                                                <Setter TargetName="Bd" Property="BorderBrush" Value="#A0EEC0"/>
                                            </Trigger>
                                            <Trigger Property="IsPressed" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="#4469F0AE"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Button.Template>
                                <TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="16" Foreground="#69F0AE"/>
                            </Button>
                        </Grid>
                    </Border>
"@
        }
    }

    return $taskItemsXaml
}

function Show-TaskWindow {
    param(
        [string]$FilePath,
        [hashtable]$SavedConfig
    )

    # Capture cache reference before any closures (see MEMORY.md for scoping notes)
    $headersCache = $script:wsHeadersCache

    # get worksheet numbers from config memory for purposes of loading bar or default to 1 if there arent any
    $loadingMax = if ($SavedConfig["WORKSHEET_NAMES"]) { (@($SavedConfig["WORKSHEET_NAMES"] -split ',' | Where-Object { $_ }).Count) + 1} else { 1 }

    [xml]$xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    Title="TaskMonitor"
    Icon="$icon"
    SizeToContent="Height" Width="620"
    MaxHeight="820"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    ShowTitleBar="False"
    UseNoneWindowStyle="True"
    MinHeight="250" MinWidth="450"
    Background="#1E1E2E"
    GlowBrush="#00BCD4"
    NonActiveGlowBrush="#333333">

    $(Get-WindowResourcesXaml)

    <Window.TaskbarItemInfo>
        <TaskbarItemInfo x:Name="TaskbarInfo"/>
    </Window.TaskbarItemInfo>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        $(Get-TitleBarXaml -Title 'TaskMonitor' -Buttons @(
            @{ Name='SettingsBtn';      Icon='&#xE713;'; Style='TitleBarBtn';      Tooltip='Open Settings' }
            @{ Name='OpenSheetBtn';     Icon='&#xE8E5;'; Style='TitleBarBtn';      Tooltip='Open Spreadsheet File' }
            @{ Name='RefreshBtn';       Icon='&#xE72C;'; Style='TitleBarBtn';      Tooltip='Refresh Task List' }
            @{ Name='MinBtn';           Icon='&#xE921;'; Style='TitleBarBtn';      Tooltip='Minimise' }
            @{ Name='MaxBtn';           Icon='&#xE922;'; Style='TitleBarBtn';      Tooltip='Maximise' }
            @{ Name='TitleCloseBtn';    Icon='&#xE8BB;'; Style='TitleBarCloseBtn'; Tooltip='Close' }
        ))

        <!-- Loading Panel: visible during initial load and refresh -->
        <Grid x:Name="LoadingPanel" Grid.Row="1" Visibility="Visible" MinHeight="300">
            <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center" Margin="40">
                <TextBlock Text="TaskMonitor"
                           FontSize="32" FontWeight="Light" Foreground="#E0E0E0"
                           TextAlignment="Center"/>
                <TextBlock x:Name="LoadingStatus"
                           FontSize="14" Foreground="#AAAAAA"
                           TextAlignment="Center" Margin="0,16,0,0"/>
                <ProgressBar x:Name="LoadingProgress" IsIndeterminate="False" Minimum="0" Maximum="$loadingMax" 
                            Value="0" Height="4" Foreground="#00ff22" Margin="0,16,0,0"/>
            </StackPanel>
        </Grid>

        <!-- Results Panel: visible after loading completes -->
        <Grid x:Name="ResultsPanel" Grid.Row="1" Visibility="Collapsed" Margin="20,4,20,20">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Margin="0,0,0,16">
                <TextBlock x:Name="HeaderText"
                           Text=""
                           FontSize="26" FontWeight="Light" Foreground="#E0E0E0"
                           VerticalAlignment="Center"/>
                <Image x:Name="MascotImage"
                       Source="$([uri]::new($mascot).AbsoluteUri)"
                       Height="72" HorizontalAlignment="Right" VerticalAlignment="Center"
                       RenderOptions.BitmapScalingMode="HighQuality"/>
            </Grid>
            <ScrollViewer Grid.Row="1"
                          VerticalScrollBarVisibility="Auto"
                          HorizontalScrollBarVisibility="Disabled">
                <StackPanel x:Name="TaskList" Margin="0,0,8,0"/>
            </ScrollViewer>
            <TextBlock x:Name="FactTipBlock" Grid.Row="2"
                       FontSize="12" Foreground="#555577" TextAlignment="Center"
                       TextWrapping="Wrap" Margin="0,12,0,4" FontStyle="Italic"/>
        </Grid>
    </Grid>
</Controls:MetroWindow>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)

    # Taskbar event handlers (Explorer restart + session unlock) are registered
    # below, after $taskbarInfo and $loadState are in scope — see Register-TaskbarEventHandlers.

    # Find named elements
    $loadingPanel    = $window.FindName("LoadingPanel")
    $loadingStatus   = $window.FindName("LoadingStatus")
    $loadingProgress = $window.FindName("LoadingProgress")
    if (-not $loadingProgress) { throw "FindName('LoadingProgress') returned null: check XAML x:Name" }
    $resultsPanel  = $window.FindName("ResultsPanel")
    $taskList      = $window.FindName("TaskList")
    $headerText    = $window.FindName("HeaderText")
    $mascotImage   = $window.FindName("MascotImage")
    $factTipBlock  = $window.FindName("FactTipBlock")
    $taskbarInfo   = $window.FindName("TaskbarInfo")
    $settingsBtn   = $window.FindName("SettingsBtn")
    $openSheetBtn  = $window.FindName("OpenSheetBtn")
    $refreshBtn    = $window.FindName("RefreshBtn")

    # Helper: parse task XAML and repopulate the task StackPanel
    $updateTaskList = {
        param($panel, $due, $upcoming)
        $wdStr = $loadState.SavedConfig['WORKING_DAYS']
        $workingDays = if ($wdStr) { @($wdStr -split ',' | ForEach-Object { $_.Trim() }) } else { @('Monday','Tuesday','Wednesday','Thursday','Friday') }
        $xmlStr = New-TaskItemsXaml -DueTasks $due -UpcomingTasks $upcoming -WorkingDays $workingDays
        $wpfNs  = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"'
        $temp   = [System.Windows.Markup.XamlReader]::Parse("<StackPanel $wpfNs>$xmlStr</StackPanel>")
        $panel.Children.Clear()
        $kids = [System.Windows.UIElement[]]@($temp.Children)
        $temp.Children.Clear()
        foreach ($k in $kids) { $panel.Children.Add($k) | Out-Null }

        # Wire Done button click handlers
        foreach ($child in $panel.Children) {
            if ($child -isnot [System.Windows.Controls.Border]) { continue }
            $grid = $child.Child
            if ($grid -isnot [System.Windows.Controls.Panel]) { continue }
            $btn = $grid.Children | Where-Object { $_ -is [System.Windows.Controls.Button] } | Select-Object -First 1
            if (-not $btn) { continue }
            $tag = $btn.Tag
            $task = $null
            if ($tag -match '^D_(\d+)$') { $task = $due[[int]$matches[1]] }
            elseif ($tag -match '^U_(\d+)$') { $task = $upcoming[[int]$matches[1]] }
            if (-not $task) { continue }
            $capturedTask        = $task
            $capturedFilePath    = $loadState.FilePath
            $capturedTag         = $tag
            $capturedLoadState   = $loadState
            $capturedPanel       = $panel
            $capturedHeaderText  = $headerText
            $capturedTaskbarInfo = $taskbarInfo
            $capturedUpdate      = $updateTaskList
            $capturedShow        = $showResults
            $btn.Add_Click({
                $wrote = Set-TaskComplete -FilePath $capturedFilePath -WorksheetName $capturedTask.Worksheet -RowIndex $capturedTask.RowIndex
                if (-not $wrote) { return }

                # Remove from the live in-memory list
                if ($capturedTag -match '^D_') {
                    $capturedLoadState.CurrentDue.Remove($capturedTask) | Out-Null
                } else {
                    $capturedLoadState.CurrentUpcoming.Remove($capturedTask) | Out-Null
                }

                $newDue   = $capturedLoadState.CurrentDue
                $newUpcom = $capturedLoadState.CurrentUpcoming
                $total    = $newDue.Count + $newUpcom.Count

                if ($total -gt 0) {
                    $capturedHeaderText.Text = "$total $(if ($total -eq 1) { 'task requires' } else { 'tasks require' }) your attention"
                    & $capturedUpdate $capturedPanel $newDue $newUpcom
                    if ($newDue.Count -gt 0) {
                        $capturedTaskbarInfo.Overlay     = New-TaskCountOverlay -Count $newDue.Count
                        $capturedTaskbarInfo.Description = "$($newDue.Count) task(s) due or overdue"
                    } else {
                        $capturedTaskbarInfo.Overlay     = $null
                        $capturedTaskbarInfo.Description = ""
                    }
                } else {
                    # Last task cleared — hand off to $showResults for the all-done display
                    if ($capturedLoadState.FactTimer) {
                        $capturedLoadState.FactTimer.Stop()
                        $capturedLoadState.FactTimer = $null
                    }
                    & $capturedShow $newDue $newUpcom
                }
            }.GetNewClosure())
        }
    }

    # Disable interactive buttons until initial loading completes
    $settingsBtn.IsEnabled  = $false
    $openSheetBtn.IsEnabled = $false
    $refreshBtn.IsEnabled   = $false

    # State shared between the load timer and refresh/settings handlers
    $loadState = @{
        ShowText         = $true
        Phase            = 'opening'
        WsIndex          = 0
        Loading          = $true      # true while timer or refresh is running
        WorksheetNames   = @()
        SheetDataCache   = @{}
        SavedConfig      = $SavedConfig
        FilePath         = $FilePath
        AllDueTasks      = [System.Collections.Generic.List[hashtable]]::new()
        AllUpcomingTasks = [System.Collections.Generic.List[hashtable]]::new()
        FactTimer        = $null
        CurrentDue       = $null
        CurrentUpcoming  = $null
    }

    # Shared: populate ResultsPanel with tasks or all-done message
    $showResults = {
        param($due, $upcom)
        $loadState.CurrentDue      = $due
        $loadState.CurrentUpcoming = $upcom
        $total = $due.Count + $upcom.Count
        if ($total -gt 0) {
            $mascotImage.Visibility = "Visible"
            $headerText.Foreground    = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#E0E0E0")
            $headerText.TextAlignment = [System.Windows.TextAlignment]::Left
            $headerText.Text = "$total $(if ($total -eq 1) { 'task requires' } else { 'tasks require' }) your attention"
            & $updateTaskList $taskList $due $upcom
            if ($due.Count -gt 0) {
                $taskbarInfo.Overlay     = New-TaskCountOverlay -Count $due.Count
                $taskbarInfo.Description = "$($due.Count) task(s) due or overdue"
            } else {
                $taskbarInfo.Overlay     = $null
                $taskbarInfo.Description = ""
            }
        } else {
            Show-ToastNotification -Title "All Done!" -Message "No tasks due or upcoming!" -Image $successImage
            $mascotImage.Visibility = "Collapsed"
            $headerText.Foreground    = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#338E9C")
            $headerText.TextAlignment = [System.Windows.TextAlignment]::Center
            $headerText.FontSize = 40
            $headerText.Text          = "All Done!"
            $wpfNs     = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
            $allDonePanel = [System.Windows.Markup.XamlReader]::Parse(
                "<StackPanel $wpfNs>
                   <Image Width=""150"" RenderOptions.BitmapScalingMode=""HighQuality"">
                    <Image.Source>
                        <BitmapImage DecodePixelWidth=""150"" UriSource=""$([uri]::new($mascot).AbsoluteUri)"" />
                    </Image.Source>
                   </Image>
                   <TextBlock Text=""No tasks are due today, overdue, or upcoming."" FontSize=""14"" Foreground=""#AAAAAA"" TextAlignment=""Center"" Margin=""0,8,0,0""/>
                 </StackPanel>")
            $taskList.Children.Clear()
            $taskList.Children.Add($allDonePanel) | Out-Null
            $taskbarInfo.Overlay     = $null
            $taskbarInfo.Description = ""
        }

        # Shared: start fact/tip ticker in the bottom strip (shown for both tasks and no-tasks views)
        $tips = @(
            "Tip: Set your working days in settings so TaskMonitor can warn you when a due date falls on a non-working day.",
            "Tip: Completed a task? Hit the tick button to update your spreadsheet and your list immediately.",
            "Tip: You can open your spreadsheet directly from the toolbar if you need to. No need to hunt for the file!",
            "Tip: Overdue tasks show in red, upcoming in blue. Non-working day tasks are highlighted orange, and others are highlighted yellow.",
            "Tip: The taskbar icon badge shows how many tasks are due or overdue at a glance.",
            "Tip: Upcoming tasks give you advance notice so you can plan ahead, or complete tasks early.",
            "Tip: You can sync your spreadsheet with OneDrive or another cloud service and track your tasks over multiple computers.",
            "Tip: TaskMonitor can run at Windows startup so your task list is always ready when you are. Enable in settings.",
            "Tip: TaskMonitor can be set to close/minimise to your system tray for a bit of headspace, or that `"out of sight out of mind`" approach to task monitoring. Just double-click and your right back onto your task list again."
        )
        $block = $factTipBlock  # local copy so GetNewClosure() captures it
        $showFactOrTip = {
            if ((Get-Random -Maximum 2) -eq 0) {
                $block.Text = Get-Random -InputObject $tips
            } else {
                $f = try { (Invoke-RestMethod "https://uselessfacts.jsph.pl/api/v2/facts/random").text } catch { "" }
                $block.Text = "Fun fact: $f"
            }
        }.GetNewClosure()
        & $showFactOrTip
        $factTimer = New-Object System.Windows.Threading.DispatcherTimer
        $factTimer.Interval = [TimeSpan]::FromSeconds(30)
        $factTimer.Add_Tick({ & $showFactOrTip }.GetNewClosure())
        $factTimer.Start()
        $loadState.FactTimer = $factTimer
    }

    # Refresh: show loading panel, reload all worksheets, then show results
    $refreshTasks = {
        if ($loadState.Loading) { return }
        if ($loadState.FactTimer) { $loadState.FactTimer.Stop(); $loadState.FactTimer = $null }
        $loadState.Loading       = $true
        $resultsPanel.Visibility = [System.Windows.Visibility]::Collapsed
        $loadingPanel.Visibility = [System.Windows.Visibility]::Visible
        $settingsBtn.IsEnabled   = $false
        $refreshBtn.IsEnabled    = $false
        $loadingStatus.Text      = "Refreshing..."
        $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Render, [action]{})
        $newDue      = [System.Collections.Generic.List[hashtable]]::new()
        $newUpcoming = [System.Collections.Generic.List[hashtable]]::new()
        $latestConfig = Load-Config
        foreach ($ws in $loadState.WorksheetNames) {
            $loadingStatus.Text = "Reading '$ws'..."
            $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Render, [action]{})
            try {
                $sr = Load-Spreadsheet -FilePath $loadState.FilePath -WorksheetName $ws
                if (-not $sr.Data) { continue }
                if ($sr.Data.Count -gt 0) {
                    $headersCache[$ws] = Get-SheetHeaders $sr.Data[0]
                }
                $wsEnv = Get-WsEnvName $ws
                $dc    = $latestConfig["${wsEnv}_DATE_COLUMN"]
                $dsc   = $latestConfig["${wsEnv}_DESCRIPTION_COLUMN"]
                if (-not $dc -or -not $dsc) { continue }
                $r = Get-DueTasks -Data $sr.Data -DateColumn $dc -DescriptionColumn $dsc -worksheetName $ws
                foreach ($t in $r.Due)      { $t['Worksheet'] = $ws; $newDue.Add($t) }
                foreach ($t in $r.Upcoming) { $t['Worksheet'] = $ws; $newUpcoming.Add($t) }
            } catch { continue }
        }
        $loadingPanel.Visibility = [System.Windows.Visibility]::Collapsed
        $resultsPanel.Visibility = [System.Windows.Visibility]::Visible
        $settingsBtn.IsEnabled   = $true
        $openSheetBtn.IsEnabled  = $true
        $refreshBtn.IsEnabled    = $true
        $loadState.Loading       = $false
        & $showResults $newDue $newUpcoming
    }

    # Wire up title bar drag
    $titleBar = $window.FindName("TitleBar")
    $titleBar.Add_MouseLeftButtonDown({ $window.DragMove() }.GetNewClosure())

    # Settings button
    $settingsBtn.Add_Click({
        $settingsResult = Show-SettingsWindow -FilePath $loadState.FilePath `
                                              -WorksheetNames $loadState.WorksheetNames `
                                              -SavedConfig (Load-Config) -OwnerWindow $window `
                                              -WsHeadersCache $headersCache
        if (-not $settingsResult.Confirmed) { return }
        foreach ($ws in $loadState.WorksheetNames) {
            $wsSettings = $settingsResult.WorksheetSettings[$ws]
            if ($wsSettings -and $wsSettings.DateCol -and $wsSettings.DescCol) {
                Save-Config -FilePath $settingsResult.NewFilePath -Worksheet $ws `
                           -DateColumn $wsSettings.DateCol -DescriptionColumn $wsSettings.DescCol `
                           -WorksheetNames $loadState.WorksheetNames
            }
        }
        if ($null -ne $settingsResult.WorkingDays -or $null -ne $settingsResult.MinimiseToTray) {
            $cfg = Load-Config
            if ($null -ne $settingsResult.WorkingDays)    { $cfg['WORKING_DAYS']     = $settingsResult.WorkingDays }
            if ($null -ne $settingsResult.MinimiseToTray) { $cfg['MINIMISE_TO_TRAY'] = $settingsResult.MinimiseToTray.ToString() }
            ($cfg.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) | Set-Content $configFile
            $loadState.SavedConfig = Load-Config
        }
        if ($settingsResult.NewFilePath -ne $loadState.FilePath) {
            $window.Close()
            Start-Process powershell -ArgumentList "-File `"$PSCommandPath`""
            return
        }
        & $refreshTasks
    }.GetNewClosure())

    # Open spreadsheet button
    $openSheetBtn.Add_Click({
        if ($loadState.FilePath -and (Test-Path $loadState.FilePath)) {
            Start-Process $loadState.FilePath
        }
    }.GetNewClosure())

    # Refresh button
    $refreshBtn.Add_Click({ & $refreshTasks }.GetNewClosure())

    # System tray support
    Initialize-SystemTray -Window $window -LoadState $loadState -Icon $icon | Out-Null

    # Window chrome buttons
    $minBtn = $window.FindName("MinBtn")
    $minBtn.Add_Click({
        if ($script:trayLoadState.SavedConfig['MINIMISE_TO_TRAY'] -eq 'True') {
            Hide-ToSystemTray
        } else {
            $script:trayWindow.WindowState = 'Minimized'
        }
    })
    $maxBtn = $window.FindName("MaxBtn")
    $maxBtn.Add_Click({
        if ($window.WindowState -eq 'Maximized') { $window.WindowState = 'Normal' }
        else { $window.WindowState = 'Maximized' }
    }.GetNewClosure())
    $titleCloseBtn = $window.FindName("TitleCloseBtn")
    $titleCloseBtn.Add_Click({ $window.Close() }.GetNewClosure())

    # DispatcherTimer loading state machine.
    # Double-tick pattern: ShowText=true tick updates status text (fast, no blocking);
    # ShowText=false tick does actual work. WPF renders between Background-priority ticks,
    # so status text is painted before each blocking COM operation.
    $loadTimer = New-Object System.Windows.Threading.DispatcherTimer
    $loadTimer.Interval = [TimeSpan]::FromMilliseconds(50)
    $loadTimer.Add_Tick({

        # OPENING: verify file, call Load-Spreadsheet to discover worksheets
        if ($loadState.Phase -eq 'opening') {
            if ($loadState.ShowText) {
                $loadingStatus.Text = if (-not $loadState.FilePath -or -not (Test-Path $loadState.FilePath)) {
                    "No spreadsheet configured..."
                } else {
                    "Opening spreadsheet..."
                }
                $loadingProgress.Value++
                Write-Log "[PROGRESS] Opening tick: Maximum=$($loadingProgress.Maximum) Value=$($loadingProgress.Value)"
                $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Render, [action]{})
                $loadState.ShowText = $false
            } else {
                if (-not $loadState.FilePath -or -not (Test-Path $loadState.FilePath)) {
                    $choice = Show-FileChoiceDialog
                    if (-not $choice) { $loadTimer.Stop(); $window.Close(); return }
                    $newPath = if ($choice -eq 'Create') { New-ExampleSpreadsheet } else { Select-SpreadsheetFile }
                    if (-not $newPath) { $loadTimer.Stop(); $window.Close(); return }
                    $loadState.FilePath = $newPath
                    if ($choice -eq 'Create') {
                        $exampleNames = @("Weekly Tasks", "Monthly Tasks", "Quarterly Tasks", "6-Monthly Tasks", "Annual Tasks")
                        foreach ($en in $exampleNames) {
                            Save-Config -FilePath $newPath -Worksheet $en -DateColumn "Due Date" `
                                -DescriptionColumn "Task Description" -WorksheetNames $exampleNames
                        }
                        $loadState.SavedConfig = Load-Config
                    } else {
                        $loadState.SavedConfig = @{}
                    }
                }
                $result = Load-Spreadsheet -FilePath $loadState.FilePath
                if ($result.WorksheetNames.Count -eq 0 -and $result.Data) {
                    # CSV file: synthesise a single worksheet name and cache the data
                    $loadState.WorksheetNames         = @("Main")
                    $loadState.SheetDataCache["Main"] = $result.Data
                } else {
                    $loadState.WorksheetNames = $result.WorksheetNames
                }
                $savedWsNames = @()
                if ($loadState.SavedConfig["WORKSHEET_NAMES"]) {
                    $savedWsNames = $loadState.SavedConfig["WORKSHEET_NAMES"] -split ',' |
                                    Where-Object { $_ }
                }
                $needsReconfig = $loadState.WorksheetNames.Count -gt 0 -and (
                    $savedWsNames.Count -eq 0 -or
                    $null -ne (Compare-Object $savedWsNames $loadState.WorksheetNames -EA SilentlyContinue))
                Write-Log "[PROGRESS] Opening work done: WorksheetNames=$($loadState.WorksheetNames -join ',') needsReconfig=$needsReconfig Maximum=$($loadingProgress.Maximum) Value=$($loadingProgress.Value)"
                $loadState.Phase    = if ($needsReconfig) { 'reconfig' } else { 'worksheets' }
                $loadState.ShowText = $true
            }
        }

        # RECONFIG: first-time or structure-changed setup - column selection for each worksheet
        elseif ($loadState.Phase -eq 'reconfig') {
            if ($loadState.ShowText) {
                $loadingStatus.Text = "First-time setup - please select columns for each worksheet..."
                $loadState.ShowText = $false
            } else {
                foreach ($ws in $loadState.WorksheetNames) {
                    if (-not $loadState.SheetDataCache.ContainsKey($ws)) {
                        $sr = Load-Spreadsheet -FilePath $loadState.FilePath -WorksheetName $ws
                        $loadState.SheetDataCache[$ws] = $sr.Data
                    }
                    $sheetData = $loadState.SheetDataCache[$ws]
                    if ($sheetData -and $sheetData.Count -gt 0) {
                        $headersCache[$ws] = Get-SheetHeaders $sheetData[0]
                    }
                    $col = Get-ColumnSelection -Data $sheetData -SavedConfig $loadState.SavedConfig `
                        -WorksheetName $ws -WorksheetNames $loadState.WorksheetNames
                    if ($col.SaveConfig) {
                        Save-Config -FilePath $loadState.FilePath -Worksheet $ws `
                            -DateColumn $col.DateCol -DescriptionColumn $col.DescCol `
                            -WorksheetNames $loadState.WorksheetNames
                    }
                }
                $loadState.SavedConfig = Load-Config
                $loadState.Phase       = 'worksheets'
                $loadState.ShowText    = $true
            }
        }

        # WORKSHEETS: load and extract tasks per worksheet, one at a time (text tick then work tick)
        elseif ($loadState.Phase -eq 'worksheets') {
            $wsNames = $loadState.WorksheetNames
            if ($loadState.WsIndex -lt $wsNames.Count) {
                $ws = $wsNames[$loadState.WsIndex]
                if ($loadState.ShowText) {
                    $loadingStatus.Text = "Reading '$ws'..."
                    $loadingProgress.Value++
                    Write-Log "[PROGRESS] Worksheet text tick: ws='$ws' WsIndex=$($loadState.WsIndex) Maximum=$($loadingProgress.Maximum) Value=$($loadingProgress.Value)"
                    $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Render, [action]{})
                    $loadState.ShowText = $false
                } else {
                    $tStart = Get-Date
                    try {
                        $sheetData = if ($loadState.SheetDataCache.ContainsKey($ws)) {
                            $loadState.SheetDataCache[$ws]
                        } else {
                            (Load-Spreadsheet -FilePath $loadState.FilePath -WorksheetName $ws).Data
                        }
                        if ($sheetData -and $sheetData.Count -gt 0) {
                            $headersCache[$ws] = Get-SheetHeaders $sheetData[0]
                        }
                        $sel = Get-ColumnSelection -Data $sheetData -SavedConfig $loadState.SavedConfig `
                            -WorksheetName $ws -WorksheetNames $wsNames
                        if ($sel.Confirmed) {
                            if ($sel.SaveConfig) {
                                Save-Config -FilePath $loadState.FilePath -Worksheet $ws `
                                    -DateColumn $sel.DateCol -DescriptionColumn $sel.DescCol `
                                    -WorksheetNames $wsNames
                                $loadState.SavedConfig = Load-Config
                            }
                            $r = Get-DueTasks -Data $sheetData -DateColumn $sel.DateCol `
                                              -DescriptionColumn $sel.DescCol -worksheetName $ws
                            foreach ($t in $r.Due)      { $t['Worksheet'] = $ws; $loadState.AllDueTasks.Add($t) }
                            foreach ($t in $r.Upcoming) { $t['Worksheet'] = $ws; $loadState.AllUpcomingTasks.Add($t) }
                        }
                    } catch { Write-Log "[PROGRESS] Worksheet work error on '$ws': $_" "ERROR" }
                    $elapsed = [math]::Round(((Get-Date) - $tStart).TotalMilliseconds)
                    $loadState.WsIndex++
                    Write-Log "[PROGRESS] Worksheet work done: ws='$ws' elapsed=${elapsed}ms WsIndex=$($loadState.WsIndex) Maximum=$($loadingProgress.Maximum)"
                    $loadState.ShowText = $true
                }
            } else {
                $loadState.Phase = 'done'
            }
        }

        # DONE: stop timer, transition to results view
        elseif ($loadState.Phase -eq 'done') {
            $loadTimer.Stop()
            $loadState.Loading       = $false
            $loadingPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $resultsPanel.Visibility = [System.Windows.Visibility]::Visible
            $settingsBtn.IsEnabled   = $true
            $openSheetBtn.IsEnabled  = $true
            $refreshBtn.IsEnabled    = $true
            $due   = $loadState.AllDueTasks
            $upcom = $loadState.AllUpcomingTasks
            if (($due.Count + $upcom.Count) -gt 0) {
                Show-ToastNotification -Title "Tasks due!" -Message "$($due.Count + $upcom.Count) task(s) need attention." -Image $warnImage
            }
            & $showResults $due $upcom
        }

    }.GetNewClosure())
    $loadTimer.Start()

    # Re-apply AUMID, icon, and overlay badge after Explorer restart or session unlock.
    # $taskbarInfo and $loadState are both in scope here so the closure captures them.
    $reapplyTaskbarState = {
        [Shell32.NativeMethods]::SetCurrentProcessExplicitAppUserModelID('TaskMonitor') | Out-Null
        $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)
        if ($loadState.CurrentDue -and $loadState.CurrentDue.Count -gt 0) {
            $taskbarInfo.Overlay     = New-TaskCountOverlay -Count $loadState.CurrentDue.Count
            $taskbarInfo.Description = "$($loadState.CurrentDue.Count) task(s) due or overdue"
        } else {
            $taskbarInfo.Overlay     = $null
            $taskbarInfo.Description = ''
        }
    }.GetNewClosure()

    Register-TaskbarEventHandlers -Window $window -OnReapply $reapplyTaskbarState
    $window.Add_Closed({ Unregister-TaskbarEventHandlers }.GetNewClosure())

    $window.ShowDialog() | Out-Null
}

#endregion

#region Main Logic

function Get-ColumnSelection {
    param(
        [array]$Data,
        [hashtable]$SavedConfig,
        [string]$WorksheetName,
        [string[]]$WorksheetNames
    )

    if (-not $Data -or $Data.Count -lt 2) {
        return @{ DateCol = $null; DescCol = $null; Confirmed = $true; SaveConfig = $false }
    }

    $columns = Get-SheetHeaders $Data[0]

    if ($columns.Count -eq 0) {
        $columns = @(1..$Data[0].Count | ForEach-Object { "Column_$_" })
    }

    if ($columns.Count -eq 0) {
        return @{ DateCol = $null; DescCol = $null; Confirmed = $true; SaveConfig = $false }
    }

    $worksheetNameEnv = Get-WsEnvName $WorksheetName

    $defaultDate = if ($SavedConfig["${worksheetNameEnv}_DATE_COLUMN"]) {
        $SavedConfig["${worksheetNameEnv}_DATE_COLUMN"]
    } elseif ($columns.Count -gt 2) {
        $columns[2]
    } else {
        $columns[0]
    }

    $defaultDesc = if ($SavedConfig["${worksheetNameEnv}_DESCRIPTION_COLUMN"]) {
        $SavedConfig["${worksheetNameEnv}_DESCRIPTION_COLUMN"]
    } else {
        $columns[0]
    }

    if ($columns -notcontains $defaultDate) {
        $defaultDate = if ($columns.Count -gt 2) { $columns[2] } else { $columns[0] }
    }

    if ($columns -notcontains $defaultDesc) {
        $defaultDesc = $columns[0]
    }

    $savedWorksheetNames = if ($SavedConfig["WORKSHEET_NAMES"]) {
        $SavedConfig["WORKSHEET_NAMES"] -split ','
    } else {
        @()
    }

    $useSaved = ($savedWorksheetNames.Count -gt 0) -and
                ($WorksheetNames.Count -gt 0) -and
                (Compare-Object $savedWorksheetNames $WorksheetNames).Count -eq 0 -and
                $defaultDate -and $defaultDesc

    if (-not $useSaved) {
        $result = Show-ColumnSelectionWindow -Columns $columns -WorksheetName $WorksheetName `
                                              -SavedConfig $SavedConfig -DefaultDate $defaultDate `
                                              -DefaultDesc $defaultDesc
        return $result
    }

    return @{
        DateCol = $defaultDate
        DescCol = $defaultDesc
        Confirmed = $true
        SaveConfig = $false
    }
}

function Start-TaskMonitor {
    $savedConfig = Load-Config
    if (-not $savedConfig.ContainsKey('WORKING_DAYS')) {
        $selected = Show-WorkingDaysDialog
        $savedConfig['WORKING_DAYS'] = if ($selected) { $selected } else { 'Monday,Tuesday,Wednesday,Thursday,Friday' }
        ($savedConfig.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) | Set-Content $configFile
    }
    $filePath    = $savedConfig["SPREADSHEET_PATH"]
    # Show-TaskWindow handles file selection if $filePath is missing or invalid
    Show-TaskWindow -FilePath $filePath -SavedConfig $savedConfig
}

#endregion
