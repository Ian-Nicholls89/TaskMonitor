
if (-not (Test-Path $successImage)) {
    try {
        Invoke-WebRequest -Uri "https://upload.wikimedia.org/wikipedia/commons/8/80/Checkmark_on_Circle.png" -OutFile $successImage -TimeoutSec 5
    } catch {
        $successImage = $null
    }
}
if (-not (Test-Path $warnImage)) {
    try {
        Invoke-WebRequest -Uri "https://upload.wikimedia.org/wikipedia/commons/5/55/Warningfv3.png" -OutFile $warnImage -TimeoutSec 5
    } catch {
        $warnImage = $null
    }
}

# Register AppUserModelID with the .ico path so Explorer always uses the correct taskbar icon
$regPath = "HKCU:\Software\Classes\AppUserModelId\TaskMonitor"
New-Item -Path $regPath -Force | Out-Null
New-ItemProperty -Path $regPath -Name DisplayName -Value 'TaskMonitor' -Force | Out-Null
New-ItemProperty -Path $regPath -Name IconUri     -Value $icon         -Force | Out-Null

# --- Windows Toast Notification (Enhanced) ---
function Show-ToastNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Image
    )
    try {
      [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
      [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null

      $template = @"
      <toast>
        <visual>
          <binding template="ToastGeneric">
            <image placement="appLogoOverride" src="$Image"/>
            <text>$Title</text>
            <text>$Message</text>
          </binding>
        </visual>
      </toast>
"@

      $xml = [Windows.Data.Xml.Dom.XmlDocument]::new()
      $xml.LoadXml($template)

      $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)

      [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('TaskMonitor').Show($toast)
    } catch {
        # Toast notifications are non-critical; silently ignore failures
    }
}

$t0 = Get-Date

# App data dir, config file, and debug log set up early so logging is available during init
$appDataDir = if ($env:APPDATA) { Join-Path $env:APPDATA "TaskMonitor" } else { Join-Path $env:USERPROFILE ".TaskMonitor" }
if (-not (Test-Path $appDataDir)) { New-Item -ItemType Directory -Path $appDataDir -Force | Out-Null }
$configFile = Join-Path $appDataDir "config.ini"
$logFile    = Join-Path $appDataDir "debug.log"
"" | Set-Content $logFile -Encoding UTF8

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "HH:mm:ss.fff"
    "[$ts][$Level] $Message" | Add-Content $logFile -Encoding UTF8
}

Write-Log "TaskMonitor Script (WPF). Made by Ian Nicholls, 2025."
Write-Log "Starting task checker script..."

function Write-Status {
    param([string]$Message, [string]$Status)
    $level = switch ($Status) {
        "OK"      { "INFO" }
        "MISSING" { "WARN" }
        "ERROR"   { "ERROR" }
        default   { "INFO" }
    }
    Write-Log ("[{0}] {1}" -f $Status.PadRight(7), $Message) $level
}

function Get-DllPath {
    param(
        [string]$LibPath,
        [string]$PackagePattern,
        [string]$DllName
    )
    $netPattern = if ($PSVersionTable.PSVersion.Major -ge 7) { "net[6-9]" } else { "net4" }
    $dll = Get-ChildItem "$LibPath\$PackagePattern" -Recurse -Filter $DllName -ErrorAction SilentlyContinue |
           Where-Object  { $_.FullName -match $netPattern } |
           Sort-Object   { $_.FullName } |
           Select-Object -Last 1
    if (-not $dll) {
        $dll = Get-ChildItem "$LibPath\$PackagePattern" -Recurse -Filter $DllName -ErrorAction SilentlyContinue |
               Select-Object -Last 1
    }
    return $dll.FullName
}

function Initialize-WPFLibraries {
    $libPath   = "$env:USERPROFILE\PSLibs"
    $nugetPath = "$libPath\nuget.exe"

    $requiredPackages = @{
        "MahApps.Metro"         = "MahApps.Metro.dll"
        "MaterialDesignColors"  = "MaterialDesignColors.dll"
        "MaterialDesignThemes"  = "MaterialDesignThemes.Wpf.dll"
    }

    if (-not (Test-Path $libPath)) {
        New-Item -ItemType Directory -Force -Path $libPath | Out-Null
    }

    $missing = @()
    foreach ($package in $requiredPackages.GetEnumerator()) {
        $dll = Get-ChildItem "$libPath\$($package.Key)*" -Recurse -Filter $package.Value -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($dll) {
            Write-Status "$($package.Value) found" "OK"
        } else {
            Write-Status "$($package.Value) not found" "MISSING"
            $missing += $package.Key
        }
    }

    if ($missing.Count -gt 0) {
        Write-Log "Downloading $($missing.Count) missing package(s)..."

        if (-not (Test-Path $nugetPath)) {
            Write-Log "Downloading NuGet CLI..."
            try {
                Invoke-WebRequest "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetPath
                Write-Status "NuGet CLI downloaded" "OK"
            } catch {
                Write-Status "Failed to download NuGet CLI: $_" "ERROR"
                exit 1
            }
        }

        foreach ($package in $missing) {
            Write-Log "Installing $package..."
            try {
                & $nugetPath install $package -OutputDirectory $libPath -NonInteractive | Out-Null
                Write-Status "$package installed" "OK"
            } catch {
                Write-Status "Failed to install $package : $_" "ERROR"
            }
        }
    } else {
        Write-Log "All WPF libraries are present."
    }

    Write-Log "Resolving DLL paths (PowerShell v$($PSVersionTable.PSVersion.Major))..."

    $global:WPFLibs = @{
        MahApps  = Get-DllPath $libPath "MahApps.Metro*"        "MahApps.Metro.dll"
        MDColors = Get-DllPath $libPath "MaterialDesignColors*"  "MaterialDesignColors.dll"
        MDThemes = Get-DllPath $libPath "MaterialDesignThemes*"  "MaterialDesignThemes.Wpf.dll"
    }

    $allResolved = $true
    foreach ($entry in $global:WPFLibs.GetEnumerator()) {
        if ($entry.Value) {
            Write-Status "$($entry.Key): $($entry.Value)" "OK"
        } else {
            Write-Status "$($entry.Key): Could not resolve path" "ERROR"
            $allResolved = $false
        }
    }

    if (-not $allResolved) {
        Write-Log "One or more WPF DLLs could not be resolved." "ERROR"
        exit 1
    }

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $controlzEx    = Get-DllPath $libPath "ControlzEx*" "ControlzEx.dll"
    $xamlBehaviors = Get-DllPath $libPath "Microsoft.Xaml.Behaviors.Wpf*" "Microsoft.Xaml.Behaviors.dll"

    if ($controlzEx)    { [System.Reflection.Assembly]::LoadFrom($controlzEx)    | Out-Null }
    if ($xamlBehaviors) { [System.Reflection.Assembly]::LoadFrom($xamlBehaviors) | Out-Null }

    [System.Reflection.Assembly]::LoadFrom($global:WPFLibs["MDColors"]) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($global:WPFLibs["MDThemes"]) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($global:WPFLibs["MahApps"])  | Out-Null

    Write-Status "WPF libraries loaded" "OK"
}

# Load WPF libraries
Initialize-WPFLibraries

# Create a WPF Application instance with OnExplicitShutdown so Hide() does not terminate the process.
# Must be created after WPF assemblies are loaded and before any Window is shown.
if (-not [System.Windows.Application]::Current) {
    $script:wpfApp = New-Object System.Windows.Application
    $script:wpfApp.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
}

# Also load WinForms for file dialog and VisualBasic for input
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Set AppUserModelID so the taskbar shows TaskMonitor's icon instead of PowerShell's
Add-Type -MemberDefinition @'
[DllImport("shell32.dll", SetLastError=true)]
public static extern int SetCurrentProcessExplicitAppUserModelID(
    [MarshalAs(UnmanagedType.LPWStr)] string AppID);
[DllImport("user32.dll", SetLastError=true)]
public static extern uint RegisterWindowMessage(
    [MarshalAs(UnmanagedType.LPWStr)] string lpString);
'@ -Namespace Shell32 -Name NativeMethods
[Shell32.NativeMethods]::SetCurrentProcessExplicitAppUserModelID("TaskMonitor") | Out-Null

$t1 = Get-Date
$loadTime = ($t1 - $t0).TotalSeconds
Write-Log "Module load complete in $([math]::Round($loadTime, 2))s"

# Ensure config file exists
if (-not (Test-Path $configFile)) {
    New-Item -ItemType File -Path $configFile -Force | Out-Null
}
Write-Log "Script settings located at $appDataDir"
