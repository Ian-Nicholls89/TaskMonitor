
function Show-WorkingDaysDialog {
    [xml]$xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    Title="TaskMonitor"
    Icon="$icon"
    Width="400" SizeToContent="Height"
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
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        $(Get-TitleBarXaml -Title 'TaskMonitor' -Buttons @(
            @{ Name='TitleCloseBtn'; Icon='&#xE8BB;'; Style='TitleBarCloseBtn' }
        ))

        <StackPanel Grid.Row="1" Margin="28,16,28,28">
            <TextBlock Text="Working days"
                       FontSize="22" FontWeight="Light" Foreground="#E0E0E0" Margin="0,0,0,6"/>
            <TextBlock Text="Select the days you work. Upcoming tasks due on non-working days will be highlighted in amber."
                       FontSize="13" Foreground="#888888" TextWrapping="Wrap" Margin="0,0,0,20"/>
            <UniformGrid Columns="7" Margin="0,0,0,24">
                <CheckBox x:Name="ChkMonday"    Content="Mon" IsChecked="True"  Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkTuesday"   Content="Tue" IsChecked="True"  Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkWednesday" Content="Wed" IsChecked="True"  Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkThursday"  Content="Thu" IsChecked="True"  Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkFriday"    Content="Fri" IsChecked="True"  Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkSaturday"  Content="Sat" IsChecked="False" Foreground="#E0E0E0" Margin="0,0,6,0"/>
                <CheckBox x:Name="ChkSunday"    Content="Sun" IsChecked="False" Foreground="#E0E0E0"/>
            </UniformGrid>
            <Button x:Name="SaveBtn"
                    Style="{StaticResource MaterialDesignRaisedButton}"
                    Background="#00BCD4" BorderBrush="#00BCD4" Foreground="White"
                    FontSize="14" Height="44"
                    materialDesign:ButtonAssist.CornerRadius="6"
                    Content="Save"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg    = [Windows.Markup.XamlReader]::Load($reader)
    $dlg.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)

    $titleBar = $dlg.FindName("TitleBar")
    $titleBar.Add_MouseLeftButtonDown({ $dlg.DragMove() }.GetNewClosure())
    $dlg.FindName("TitleCloseBtn").Add_Click({ $dlg.Close() }.GetNewClosure())

    $state = @{ Days = $null; Dlg = $dlg }
    $dlg.FindName("SaveBtn").Add_Click({
        $checked = @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday') | Where-Object {
            $chk = $state.Dlg.FindName("Chk$_")
            $chk -and $chk.IsChecked
        }
        $state.Days = if ($checked) { $checked -join ',' } else { 'Monday,Tuesday,Wednesday,Thursday,Friday' }
        $state.Dlg.Close()
    }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
    return $state.Days
}
