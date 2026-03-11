
#region XAML Builder Helpers

function Get-WindowResourcesXaml {
    # Single definition of all WPF theme merges + both title-bar button styles.
    # Embed in any MetroWindow XAML via $(Get-WindowResourcesXaml).
    return @'
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Teal.xaml"/>
                <materialDesign:BundledTheme BaseTheme="Dark" PrimaryColor="Teal" SecondaryColor="Cyan"/>
                <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.Wpf;component/Themes/MaterialDesign3.Defaults.xaml"/>
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="TitleBarBtn" TargetType="Button">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="#999999"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Width" Value="46"/>
                <Setter Property="Height" Value="36"/>
                <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
                <Setter Property="FontSize" Value="10"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="{TemplateBinding Background}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#2A2A3E"/>
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#3A3A4E"/>
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="TitleBarCloseBtn" TargetType="Button">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="#999999"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Width" Value="46"/>
                <Setter Property="Height" Value="36"/>
                <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
                <Setter Property="FontSize" Value="10"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border x:Name="Bd" Background="{TemplateBinding Background}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#E81123"/>
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="Bd" Property="Background" Value="#C50F1F"/>
                                    <Setter Property="Foreground" Value="#FFFFFF"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
'@
}

function Get-SheetHeaders {
    # Extracts non-empty, trimmed header values from the first row of sheet data.
    param([array]$Row)
    return @($Row | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString().Trim() })
}

function Get-WsEnvName {
    # Returns the config-key prefix for a worksheet name (e.g. "My Sheet" -> "MY_SHEET").
    param([string]$WorksheetName)
    return $WorksheetName.ToUpper() -replace '\s+', '_'
}

function Get-ComboItemsXaml {
    # Builds a XAML string of <ComboBoxItem> elements for the given column names.
    param([string[]]$Columns)
    $items = ""
    foreach ($col in $Columns) {
        $esc = [System.Security.SecurityElement]::Escape($col)
        $items += "                        <ComboBoxItem Content=`"$esc`"/>`n"
    }
    return $items
}

function Get-TitleBarXaml {
    param(
        [string]$Title,
        [array]$Buttons   # each: @{ Name='...'; Icon='&#x...;'; Style='TitleBarBtn'|'TitleBarCloseBtn' }
    )
    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $colDefs  = "            <ColumnDefinition Width=`"*`"/>`n"
    $btnElems = ""
    for ($i = 0; $i -lt $Buttons.Count; $i++) {
        $b   = $Buttons[$i]
        $col = $i + 1
        $colDefs  += "            <ColumnDefinition Width=`"Auto`"/>`n"
        $tooltip   = if ($b.Tooltip) { " ToolTip=`"$($b.Tooltip)`"" } else { "" }
        $btnElems += "            <Button x:Name=`"$($b.Name)`" Grid.Column=`"$col`" Content=`"$($b.Icon)`" Style=`"{StaticResource $($b.Style)}`"$tooltip/>`n"
    }
    return @"
        <Grid x:Name="TitleBar" Grid.Row="0" Background="#1E1E2E" Height="36">
            <Grid.ColumnDefinitions>
$colDefs            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="$escapedTitle"
                       FontSize="12" Foreground="#999999"
                       VerticalAlignment="Center" Margin="12,0,0,0"/>
$btnElems        </Grid>
"@
}

#endregion
