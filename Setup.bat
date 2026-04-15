@echo off
@title Media Scripts 2026 - AIO
cd /d "%~dp0"

echo ================================================================
echo  Media Scripts 2026 - AIO
echo  Creates folders, writes bat files, and downloads all tools.
echo ================================================================
echo.
echo  Options (leave blank for full setup):
echo    --force             Re-download tools and overwrite bat files
echo    --directories-only  Create folders only, skip everything else
echo    --skip-bat-files    Skip writing bat files
echo    --skip-handbrake    Skip HandBrake (~65 MB)
echo    --skip-staxrip      Skip StaxRip
echo    --skip-tsmuxer      Skip tsMuxer / tsMuxerGUI
echo.
echo  To compile to exe (requires pyinstaller):
echo    pip install pyinstaller
echo    pyinstaller --onefile --name media_scripts_setup setup.py
echo.
set /p ARGS="Extra flags (or press Enter for full setup): "

REM --- Try frozen exe first, fall back to python ---
if exist "%~dp0media_scripts_setup.exe" (
    "%~dp0media_scripts_setup.exe" %ARGS%
) else (
    python "%~dp0setup.py" %ARGS%
    if errorlevel 1 (
        echo.
        echo ERROR: Python not found or setup.py failed.
        echo Make sure Python 3.10+ is installed and on PATH.
        echo Or compile setup.py to media_scripts_setup.exe with build-exe.bat.
    )
)

echo.
pause
