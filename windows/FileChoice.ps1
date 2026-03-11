
function Show-FileChoiceDialog {
    [xml]$xaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    Title="TaskMonitor"
    Icon="$icon"
    Width="420" SizeToContent="Height"
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
            <TextBlock Text="Get started"
                       FontSize="22" FontWeight="Light" Foreground="#E0E0E0" Margin="0,0,0,6"/>
            <TextBlock Text="No spreadsheet has been configured yet."
                       FontSize="13" Foreground="#888888" Margin="0,0,0,24"/>
            <Button x:Name="LoadBtn"
                    Style="{StaticResource MaterialDesignRaisedButton}"
                    Background="#00BCD4" BorderBrush="#00BCD4" Foreground="White"
                    FontSize="14" Height="44" Margin="0,0,0,12"
                    materialDesign:ButtonAssist.CornerRadius="6">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE8B7;" FontFamily="Segoe MDL2 Assets" FontSize="16" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Text="Load existing spreadsheet" VerticalAlignment="Center"/>
                </StackPanel>
            </Button>
            <Button x:Name="CreateBtn"
                    Style="{StaticResource MaterialDesignOutlinedButton}"
                    Foreground="#69F0AE" BorderBrush="#69F0AE"
                    FontSize="14" Height="44"
                    materialDesign:ButtonAssist.CornerRadius="6">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="&#xE8A5;" FontFamily="Segoe MDL2 Assets" FontSize="16" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Text="Create example spreadsheet" VerticalAlignment="Center"/>
                </StackPanel>
            </Button>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    $dlg.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$icon)

    $titleBar = $dlg.FindName("TitleBar")
    $titleBar.Add_MouseLeftButtonDown({ $dlg.DragMove() }.GetNewClosure())
    $dlg.FindName("TitleCloseBtn").Add_Click({ $dlg.Close() }.GetNewClosure())

    $state = @{ Choice = $null }
    $dlg.FindName("LoadBtn").Add_Click({   $state.Choice = 'Load';   $dlg.Close() }.GetNewClosure())
    $dlg.FindName("CreateBtn").Add_Click({ $state.Choice = 'Create'; $dlg.Close() }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
    return $state.Choice
}
