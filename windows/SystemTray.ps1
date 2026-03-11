
#region System Tray

function Initialize-SystemTray {
    param(
        [System.Windows.Window]$Window,
        [hashtable]$LoadState,
        [string]$Icon
    )

    # Store in script scope so all event handlers share the same state without GetNewClosure
    $script:trayWindow      = $Window
    $script:trayLoadState   = $LoadState
    $script:trayForceClose  = $false
    $script:trayRestoreLeft = $null
    $script:trayRestoreTop  = $null

    $script:trayNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    try {
        $script:trayNotifyIcon.Icon = [System.Drawing.Icon]::new($Icon)
    } catch {
        Write-Log "NotifyIcon: icon load failed: $_" "ERROR"
    }
    $script:trayNotifyIcon.Text    = "TaskMonitor"
    $script:trayNotifyIcon.Visible = $false

    $trayMenu     = New-Object System.Windows.Forms.ContextMenu
    $trayShowItem = New-Object System.Windows.Forms.MenuItem "Show TaskMonitor"
    $trayExitItem = New-Object System.Windows.Forms.MenuItem "Exit"
    $trayMenu.MenuItems.Add($trayShowItem) | Out-Null
    $trayMenu.MenuItems.Add($trayExitItem) | Out-Null
    $script:trayNotifyIcon.ContextMenu = $trayMenu

    $script:trayNotifyIcon.add_DoubleClick({ Restore-FromTray })
    $trayShowItem.add_Click({ Restore-FromTray })
    $trayExitItem.add_Click({
        $script:trayForceClose           = $true
        $script:trayNotifyIcon.Visible   = $false
        $script:trayNotifyIcon.Dispose()
        $script:trayWindow.ShowInTaskbar = $true
        $script:trayWindow.Close()
    })

    # Intercept window close: hide to tray if setting is on (unless exiting via tray menu)
    $Window.Add_Closing({
        param($s, $e)
        if (-not $script:trayForceClose -and $script:trayLoadState.SavedConfig['MINIMISE_TO_TRAY'] -eq 'True') {
            $e.Cancel = $true
            Hide-ToSystemTray
        }
    })

    # Clean up NotifyIcon when the window actually closes
    $Window.Add_Closed({
        if ($script:trayNotifyIcon) {
            $script:trayNotifyIcon.Visible = $false
            $script:trayNotifyIcon.Dispose()
        }
    })

    # Intercept minimise via Win+M, taskbar right-click, etc. (MinBtn handled in MainWindow)
    # Restore to Normal first so we can move off-screen cleanly (Minimized position is -32000,-32000)
    $Window.Add_StateChanged({
        if ($script:trayWindow.WindowState -eq 'Minimized' -and
            $script:trayLoadState.SavedConfig['MINIMISE_TO_TRAY'] -eq 'True') {
            $script:trayWindow.WindowState = 'Normal'
            Hide-ToSystemTray
        }
    })

    # Return a scriptblock wrapper so MainWindow can call it the same way as before
    return @{ HideToTray = { Hide-ToSystemTray } }
}

function Hide-ToSystemTray {
    Write-Log "Hide-ToSystemTray: Left=$($script:trayWindow.Left) Top=$($script:trayWindow.Top)"
    # Capture position while Normal (must be called before moving off-screen)
    $script:trayRestoreLeft          = $script:trayWindow.Left
    $script:trayRestoreTop           = $script:trayWindow.Top
    $script:trayNotifyIcon.Visible   = $true
    $script:trayWindow.ShowInTaskbar = $false
    # Move off-screen instead of minimizing — avoids taskbar ghost and keeps ShowDialog() alive
    $script:trayWindow.Left          = -32000
    $script:trayWindow.Top           = -32000
    Write-Log "Hide-ToSystemTray: done"
}

function Restore-FromTray {
    Write-Log "Restore-FromTray: Left=$($script:trayRestoreLeft) Top=$($script:trayRestoreTop)"
    $script:trayNotifyIcon.Visible   = $false
    # Restore position before re-showing in taskbar to avoid flash at off-screen coords
    if ($null -ne $script:trayRestoreLeft) { $script:trayWindow.Left = $script:trayRestoreLeft }
    if ($null -ne $script:trayRestoreTop)  { $script:trayWindow.Top  = $script:trayRestoreTop  }
    $script:trayWindow.ShowInTaskbar = $true
    $script:trayWindow.Activate()
    Write-Log "Restore-FromTray: done"
}

#endregion
