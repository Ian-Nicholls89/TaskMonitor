
#region WPF Column Selection Window

function Show-ColumnSelectionWindow {
    param(
        [string[]]$Columns,
        [string]$WorksheetName,
        [hashtable]$SavedConfig,
        [string]$DefaultDate,
        [string]$DefaultDesc
    )

    $title = if ($WorksheetName) { "Select Columns - $WorksheetName" } else { "Select Columns" }

    # Build ComboBox items XML
    $dateItems = Get-ComboItemsXaml -Columns $Columns
    $descItems = $dateItems

    # Build column list items
    $columnListItems = ""
    foreach ($col in $Columns) {
        $escaped = [System.Security.SecurityElement]::Escape($col)
        $columnListItems += "                        <TextBlock Text=`"  $escaped`" Foreground=`"#AAAAAA`" FontSize=`"13`" Margin=`"0,2`"/>`n"
    }

    $savedNotice = ""
    if ($SavedConfig.Count -gt 0) {
        $savedNotice = @"
                <TextBlock Text="Using saved configuration. Verify columns:"
                           FontSize="12" FontStyle="Italic" Foreground="#64B5F6" Margin="0,0,0,10"/>
"@
    }

    [xml]$xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    Title="$([System.Security.SecurityElement]::Escape($title))"
    Icon="$icon"
    Height="580" Width="500"
    WindowStartupLocation="CenterScreen"
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

        $(Get-TitleBarXaml -Title $title -Buttons @(
            @{ Name='TitleCloseBtn'; Icon='&#xE8BB;'; Style='TitleBarCloseBtn' }
        ))

        <Grid Grid.Row="1" Margin="24,8,24,24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0"
                   Text="Configure Columns"
                   FontSize="22" FontWeight="Light" Foreground="#E0E0E0"
                   Margin="0,0,0,6"/>

        <TextBlock Grid.Row="1"
                   Text="$([System.Security.SecurityElement]::Escape($WorksheetName))"
                   FontSize="14" Foreground="#00BCD4"
                   Margin="0,0,0,12"/>

        <!-- Saved config notice -->
        <StackPanel Grid.Row="2">
            $savedNotice
        </StackPanel>

        <!-- Available columns -->
        <StackPanel Grid.Row="3">
            <TextBlock Text="Available columns:" FontSize="13" FontWeight="Bold" Foreground="#B0B0B0" Margin="0,0,0,6"/>
            <ScrollViewer MaxHeight="140" VerticalScrollBarVisibility="Auto">
                <StackPanel>
$columnListItems
                </StackPanel>
            </ScrollViewer>
        </StackPanel>

        <!-- Date column -->
        <StackPanel Grid.Row="4" Margin="0,16,0,0">
            <TextBlock Text="DUE BY column:" FontSize="14" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,0,0,8"/>
            <ComboBox x:Name="DateCombo"
                      Style="{StaticResource MaterialDesignOutlinedComboBox}"
                      materialDesign:HintAssist.Hint="Select date column"
                      FontSize="14" Height="44">
$dateItems
            </ComboBox>
        </StackPanel>

        <!-- Description column -->
        <StackPanel Grid.Row="5" Margin="0,16,0,0">
            <TextBlock Text="TASK DESCRIPTION column:" FontSize="14" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,0,0,8"/>
            <ComboBox x:Name="DescCombo"
                      Style="{StaticResource MaterialDesignOutlinedComboBox}"
                      materialDesign:HintAssist.Hint="Select description column"
                      FontSize="14" Height="44">
$descItems
            </ComboBox>
        </StackPanel>

        <!-- Save checkbox -->
        <CheckBox x:Name="SaveCheck" Grid.Row="6"
                  Content="Save these settings for next time"
                  IsChecked="True"
                  Foreground="#CCCCCC" FontSize="13"
                  Margin="0,20,0,0"/>

        <!-- Buttons -->
        <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,20,0,0">
            <Button x:Name="ConfirmBtn"
                    Content="Confirm"
                    Style="{StaticResource MaterialDesignRaisedButton}"
                    Background="#00BCD4" BorderBrush="#00BCD4" Foreground="White"
                    FontSize="14" Width="140" Height="40"
                    materialDesign:ButtonAssist.CornerRadius="6"
                    Margin="0,0,16,0"/>
            <Button x:Name="CancelBtn"
                    Content="Cancel"
                    Style="{StaticResource MaterialDesignOutlinedButton}"
                    Foreground="#EF5350" BorderBrush="#EF5350"
                    FontSize="14" Width="140" Height="40"
                    materialDesign:ButtonAssist.CornerRadius="6"/>
        </StackPanel>
    </Grid>
    </Grid>
</Controls:MetroWindow>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)

    # Wire up custom title bar
    $titleBar = $window.FindName("TitleBar")
    $titleBar.Add_MouseLeftButtonDown({ $window.DragMove() }.GetNewClosure())
    $titleCloseBtn = $window.FindName("TitleCloseBtn")
    $titleCloseBtn.Add_Click({ $window.Close() }.GetNewClosure())

    $dateCombo  = $window.FindName("DateCombo")
    $descCombo  = $window.FindName("DescCombo")
    $saveCheck  = $window.FindName("SaveCheck")
    $confirmBtn = $window.FindName("ConfirmBtn")
    $cancelBtn  = $window.FindName("CancelBtn")

    # Set defaults
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        if ($Columns[$i] -eq $DefaultDate) { $dateCombo.SelectedIndex = $i }
        if ($Columns[$i] -eq $DefaultDesc) { $descCombo.SelectedIndex = $i }
    }

    $result = @{ Confirmed = $false }

    $confirmBtn.Add_Click({
        $result.DateCol    = $Columns[$dateCombo.SelectedIndex]
        $result.DescCol    = $Columns[$descCombo.SelectedIndex]
        $result.Confirmed  = $true
        $result.SaveConfig = $saveCheck.IsChecked
        $window.Close()
    }.GetNewClosure())

    $cancelBtn.Add_Click({
        $result.Confirmed = $false
        $window.Close()
    }.GetNewClosure())

    $window.ShowDialog() | Out-Null
    return $result
}

#endregion
