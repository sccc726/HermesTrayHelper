@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%HermesHelper.ps1"
set "VBS_PATH=%SCRIPT_DIR%Start-HermesHelper.vbs"

if not exist "%SCRIPT_PATH%" (
    echo HermesHelper.ps1 was not found next to this BAT file.
    echo Expected: "%SCRIPT_PATH%"
    pause
    exit /b 1
)

if /I "%~1"=="selftest" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -SelfTest
    exit /b %ERRORLEVEL%
)

if /I "%~1"=="debug" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
    exit /b %ERRORLEVEL%
)

if exist "%VBS_PATH%" (
    start "Hermes Helper" wscript.exe //nologo "%VBS_PATH%"
    exit /b 0
)

start "Hermes Helper" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_PATH%"
exit /b 0
