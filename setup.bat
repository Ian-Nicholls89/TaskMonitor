@echo off
setlocal enabledelayedexpansion

:: Set variable
set SCRIPT_DIR=%~dp0
set PY_MOD_SCRIPT=Setup/ModuleSetup.py
set PY_TASK_SCRIPT=TaskMonitor.py
set SCRIPT=%SCRIPT_DIR%%PY_TASK_SCRIPT%
set answer= ""

:options
echo Select setup option from the following:
echo 1. Full Setup
echo.
echo 2. Python Setup
echo 3. Python module setup only
echo.
echo 4. Activate script on Windows Startup
echo 5. Remove script from Windows Startup
echo.
echo 6. Edit script settings
echo.
CHOICE /C 123456 /M "Select Setup option"
if errorlevel 6 goto editsettings
if errorlevel 5 goto startupremove
if errorlevel 4 (
    set setup=4
    goto startupsetup)
if errorlevel 3 (
    set setup=3
    goto modulesetup)
if errorlevel 2 (
    set setup=2
    goto fullsetup)
if errorlevel 1 goto fullsetup

:fullsetup
echo Installing Python...
powershell -Command "winget install '9PJPW5LDXLZ5' -s msstore"
echo Once Python is installed press any key to continue...
pause >nul

:modulesetup
echo Running Python module setup...
python %PY_MOD_SCRIPT%

if "%setup%"=="2" goto end  
if "%setup%"=="3" goto end

:startupsetup
CHOICE /C YN /M "Do you want to run the script at Windows Startup?"
if errorlevel 2 echo Skipping Windows Startup setup. Rerun this script if you wish to run at startup in future.
if errorlevel 1 (
    echo.
    echo Setting to run on Windows Startup...
    powershell -Command "Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'TaskMonitorPythonScript' -Value 'cmd /c \"python \"%SCRIPT%\" --startup\"'"
    echo Done.
)
if "%setup%"=="4" goto end

:ask2
if not "%setup%"=="6" echo Setup done! 
echo.
CHOICE /C YN /M "Do you want to run the script now?"
echo.
if errorlevel 2 goto end
if errorlevel 1 (
    python %PY_TASK_SCRIPT%
    goto end
)

:startupremove
:: Registry paths to check
set HKCU_PATH=HKCU\Software\Microsoft\Windows\CurrentVersion\Run

:: First, show what entries exist
echo Checking for existing startup entries...
echo.

:: Check HKCU
reg query "%HKCU_PATH%" /v "TaskMonitorPythonScript" >nul 2>&1
if %errorlevel%==0 (
    echo Found script in Startup
    echo.
    :ask_removal
    CHOICE /C YN /M "Do you want to remove the startup entries?"
    echo.
    if errorlevel 2 goto end
    if errorlevel 1 (
        echo Removing startup entries...
        echo.
        reg delete "%HKCU_PATH%" /v "TaskMonitorPythonScript" /f
        goto end
    )

) else (
    echo No startup entry found
    echo Nothing to remove.
    goto end
)

:editsettings
python %PY_TASK_SCRIPT% --editsettings
goto end

:end
CHOICE /C YN /M "Do you want to setup anything else?"
if errorlevel 2 goto eof
if errorlevel 1 goto options