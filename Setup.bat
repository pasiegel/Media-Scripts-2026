@echo off
@title Media Scripts 2026 - Setup
cd /d "%~dp0"

echo ================================================================
echo  Media Scripts 2026 - Setup
echo  This will create all folders and download required tools.
echo ================================================================
echo.
echo  Options (leave blank for full setup):
echo    -Force          Re-download all tools even if already installed
echo    -DirectoriesOnly  Create folders only, skip downloads
echo    -SkipHandBrake  Skip HandBrake (~65 MB)
echo    -SkipStaxRip    Skip StaxRip
echo.
set /p ARGS="Extra flags (or press Enter): "

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup.ps1" %ARGS%

echo.
pause
