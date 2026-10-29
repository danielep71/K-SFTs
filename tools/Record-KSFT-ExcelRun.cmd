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

rem Without a dropped workbook, a fresh development workbook is built from
rem the template in the ignored test-results folder, replaced on every run.
rem A dropped workbook is synchronized as it is and never rebuilt.
set "WORKBOOK=%~1"
set "BUILD="
if not defined WORKBOOK (
    set "WORKBOOK=%CD%\test-results\K-SFT_Development.xlsm"
    set "BUILD=-FromTemplate"
)

echo K-SFT Excel run
echo.

if not defined BUILD if not exist "%WORKBOOK%" (
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
if defined BUILD (
    echo Building the workbook from src\workbook\SFT_Template.xlsx,
) else (
    echo Using the dropped workbook as it is: its sheets are not rebuilt.
)
echo synchronizing in Development mode, running the tests and checking the run...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" -RecordRun %BUILD%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: the Excel run is recorded; copy the report above.
    echo For the checks by hand, open:
    echo   %WORKBOOK%
) else if "%RC%"=="2" (
    echo NOT CHANGED: nothing was imported; the workbook was not changed.
) else if "%RC%"=="3" (
    echo NOT RECORDED: the workbook is synchronized, but the run did not pass the check.
) else (
    echo FAILED: the synchronization or the run returned exit code %RC%.
)
echo.
pause
exit /b %RC%
