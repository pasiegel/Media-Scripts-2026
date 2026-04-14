@echo off
@title Media Scripts 2026 - Setup
cd /d "%~dp0"

echo ================================================================
echo  Media Scripts 2026 - Setup
echo  Creates folders, writes bat files, downloads all tools, and
echo  patches portable Python paths so mnamer works correctly.
echo ================================================================
echo.
echo  Options (leave blank for full setup):
echo    -Force            Re-download tools and overwrite bat files
echo    -DirectoriesOnly  Create folders only, skip everything else
echo    -SkipBatFiles     Skip writing bat files
echo    -SkipHandBrake    Skip HandBrake (~65 MB)
echo    -SkipStaxRip      Skip StaxRip
echo.
echo  NOTE: If you move this folder to a new location later, run
echo        Setup_Path_Variables_If_Error.bat to re-patch Python paths.
echo.
set /p ARGS="Extra flags (or press Enter for full setup): "

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup.ps1" %ARGS%

echo.
pause
