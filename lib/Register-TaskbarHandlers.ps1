<#
.SYNOPSIS
    Registers handlers to reapply the taskbar icon and overlay after:
      - Explorer restarts (WM_TASKBARCREATED broadcast)
      - The workstation is unlocked (Microsoft.Win32.SystemEvents.SessionSwitch)

.NOTES
    The SessionSwitchEventHandler delegate fires on a background thread, so the
    callback is always marshalled onto the WPF dispatcher before executing.

    Call Register-TaskbarEventHandlers once Show-TaskWindow has its window and
    state variables in scope.  Call Unregister-TaskbarEventHandlers from the
    window's Closed event to prevent the static SystemEvents subscription from
    leaking after the window is gone.
#>

function Register-TaskbarEventHandlers {
    param(
        [System.Windows.Window]$Window,
        [scriptblock]$OnReapply   # scriptblock — called on Explorer restart OR unlock
    )

    # Store in script scope so the typed SessionSwitch delegate can reach them
    # (typed delegates do not inherit the PowerShell closure environment).
    $script:_tmReapply    = $OnReapply
    $script:_tmDispatcher = $Window.Dispatcher

    # ── 1. WM_TASKBARCREATED ────────────────────────────────────────────────
    # Explorer broadcasts this to every top-level window whenever it recreates
    # the taskbar (e.g. after Explorer.exe crashes and restarts).
    $script:_tmTaskbarMsg = [Shell32.NativeMethods]::RegisterWindowMessage('TaskbarCreated')

    $Window.Add_Loaded({
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        # Keep a script-scope reference so the HwndSource is not GC-collected.
        $script:_tmHwndSrc = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        $script:_tmHwndSrc.AddHook({
            param($hwnd, $msg, $wParam, $lParam, [ref]$handled)
            if ($msg -eq $script:_tmTaskbarMsg) {
                & $script:_tmReapply
            }
            return [IntPtr]::Zero
        })
    }.GetNewClosure())

    # ── 2. SessionSwitch (lock / unlock) ────────────────────────────────────
    # SystemEvents fires on a background thread in WPF; BeginInvoke marshals
    # the call back onto the WPF dispatcher before touching any UI state.
    $script:_tmSessionHandler = [Microsoft.Win32.SessionSwitchEventHandler] {
        param($sender, $e)
        if ($e.Reason -eq [Microsoft.Win32.SessionSwitchReason]::SessionUnlock) {
            $script:_tmDispatcher.BeginInvoke(
                [System.Windows.Threading.DispatcherPriority]::Normal,
                [Action]{ & $script:_tmReapply }
            ) | Out-Null
        }
    }
    [Microsoft.Win32.SystemEvents]::add_SessionSwitch($script:_tmSessionHandler)
}

function Unregister-TaskbarEventHandlers {
    # Must be called from the window's Closed event to remove the static
    # SystemEvents subscription (it would otherwise outlive the window).
    if ($script:_tmSessionHandler) {
        [Microsoft.Win32.SystemEvents]::remove_SessionSwitch($script:_tmSessionHandler)
    }
    $script:_tmSessionHandler = $null
    $script:_tmHwndSrc        = $null
    $script:_tmReapply        = $null
    $script:_tmDispatcher     = $null
}
