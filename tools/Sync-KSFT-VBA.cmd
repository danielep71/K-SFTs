@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Sync-KSFT-VBA.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

set "WORKBOOK=%~1"

if not defined WORKBOOK (
    echo K-SFT VBA synchronization
    echo.
    echo No workbook given. Drag the .xlsm workbook onto this file, or run:
    echo   %~nx0 ^<path to workbook^>
    echo.
    pause
    exit /b 1
)

echo K-SFT VBA synchronization
echo.

if not exist "%WORKBOOK%" (
    echo.
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

echo.
echo Repository:
echo   %CD%
echo Workbook:
echo   %WORKBOOK%
echo.
echo Synchronizing VBA source...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: K-SFT VBA synchronization completed.
) else if "%RC%"=="2" (
    echo NOT CHANGED: nothing was imported; the workbook was not changed.
) else (
    echo FAILED: K-SFT VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
