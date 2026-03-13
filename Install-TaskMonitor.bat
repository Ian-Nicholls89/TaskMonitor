@echo off
setlocal enabledelayedexpansion

set "INSTALL_DIR=%USERPROFILE%\TaskMonitor"
set "REPO_ZIP=https://github.com/Ian-Nicholls89/TaskMonitor/archive/refs/heads/main.zip"
set "ZIP_FILE=%TEMP%\TaskMonitor.zip"
set "EXTRACT_DIR=%TEMP%\TaskMonitor_extract"

echo ============================================
echo  TaskMonitor Installer
echo  Installing to: %INSTALL_DIR%
echo ============================================
echo.

if exist "%INSTALL_DIR%\TaskMonitor.ps1" (
    echo TaskMonitor is already installed at %INSTALL_DIR%.
    echo.
    choice /C YN /N /M "Do you want to uninstall? (Y/N)"
    if "!errorlevel!"=="1" (
        rmdir /S /Q "%INSTALL_DIR%"
        powershell -NoProfile -Command "$lnk = [Environment]::GetFolderPath('Desktop') + '\TaskMonitor.lnk'; if (Test-Path $lnk) { Remove-Item $lnk }"
        echo.
        echo TaskMonitor has been uninstalled.
    ) else (
        echo Uninstall cancelled.
    )
    echo.
    pause
    exit /b 0
)

echo [1/4] Downloading from GitHub...
powershell -NoProfile -Command "Invoke-WebRequest -Uri '%REPO_ZIP%' -OutFile '%ZIP_FILE%'" 2>nul
if not exist "%ZIP_FILE%" (
    echo ERROR: Download failed. Check your internet connection and try again.
    pause
    exit /b 1
)

echo [2/4] Extracting files...
if exist "%EXTRACT_DIR%" rmdir /S /Q "%EXTRACT_DIR%"
powershell -NoProfile -Command "Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%EXTRACT_DIR%' -Force"

echo [3/4] Installing to %INSTALL_DIR%...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"
xcopy /E /Y /Q "%EXTRACT_DIR%\TaskMonitor-main\*" "%INSTALL_DIR%\" >nul

echo [4/4] Cleaning up...
del "%ZIP_FILE%"
rmdir /S /Q "%EXTRACT_DIR%"

choice /C YN /N /M "Do you want to place a shortcut on your desktop? (Y/N)"
if "%errorlevel%"=="1" (
    powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\TaskMonitor.lnk'); $s.TargetPath = '%INSTALL_DIR%\TaskMonitor.vbs'; $s.IconLocation = '%INSTALL_DIR%\assets\taskmonitor.ico'; $s.Save()"
    echo Shortcut created on desktop.
)

echo.
echo ============================================
echo  Done! TaskMonitor installed to:
echo  %INSTALL_DIR%
echo.
echo  To run: double-click TaskMonitor.vbs
echo  (NuGet packages will download on first run)
echo ============================================
echo.
pause
