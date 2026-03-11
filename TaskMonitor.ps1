<#
.SYNOPSIS
    TaskMonitor robot
.DESCRIPTION
    Monitors Excel/CSV files for due tasks and displays notifications
    using MahApps.Metro + MaterialDesign WPF windows.
.NOTES
    Made by Ian Nicholls, 2025 (Python version), updated to Powershell WPF in 2026.
#>

param(
    [switch]$Startup
)

$directory    = Split-Path -Path $PSCommandPath -Parent
$successImage = Join-Path $directory "assets\tick.png"
$warnImage    = Join-Path $directory "assets\warn.png"
$mascot       = Join-Path $directory "assets\TaskMonitor\taskmonitor_512.png"
$icon         = Join-Path $directory "assets\taskmonitor.ico"

# Bypass execution policy for this process so dot-sourced lib/window scripts load without restriction
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Populated whenever a spreadsheet is loaded; used by Show-SettingsWindow to skip re-querying
$script:wsHeadersCache = @{}

. "$directory\lib\Init.ps1"
. "$directory\lib\Config.ps1"
. "$directory\lib\XamlHelpers.ps1"
. "$directory\lib\Spreadsheet.ps1"
. "$directory\lib\TaskLogic.ps1"
. "$directory\lib\Register-TaskbarHandlers.ps1"
. "$directory\windows\ColumnSelection.ps1"
. "$directory\windows\Settings.ps1"
. "$directory\windows\WorkingDays.ps1"
. "$directory\windows\FileChoice.ps1"
. "$directory\windows\SystemTray.ps1"
. "$directory\windows\MainWindow.ps1"

# Main execution
Start-TaskMonitor
