@echo off
@title Build media_scripts_setup.exe (PyInstaller)
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ================================================================
echo  Media Scripts 2026 - AIO - Build Exe
echo  Compiles setup.py into media_scripts_setup.exe using PyInstaller.
echo  Output: dist\media_scripts_setup.exe  (then moved to project root)
echo ================================================================
echo.

REM --- Check Python ---
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found on PATH.
    echo Install Python 3.10+ from https://python.org and try again.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Python: %%v

REM --- Check / install PyInstaller ---
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo PyInstaller not found.
    set /p install_pyi="Install PyInstaller now? (y/n): "
    if /i not "!install_pyi!"=="y" (
        echo Cancelled.
        pause
        exit /b 1
    )
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: pip install failed.
        pause
        exit /b 1
    )
)

for /f "tokens=*" %%v in ('python -m PyInstaller --version 2^>^&1') do echo   PyInstaller: %%v

REM --- Clean previous build artifacts ---
echo.
echo Cleaning previous build artifacts...
if exist "build"                     rmdir /s /q "build"
if exist "dist"                      rmdir /s /q "dist"
if exist "media_scripts_setup.spec"  del /q "media_scripts_setup.spec"

REM --- Build ---
echo.
echo Building media_scripts_setup.exe...
echo.
python -m PyInstaller --onefile --name media_scripts_setup --console setup.py
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed. See output above.
    pause
    exit /b 1
)

REM --- Move exe to project root ---
if exist "dist\media_scripts_setup.exe" (
    if exist "media_scripts_setup.exe" del /q "media_scripts_setup.exe"
    move "dist\media_scripts_setup.exe" "media_scripts_setup.exe" >nul
    echo.
    echo ================================================================
    echo  Build successful!
    echo  Output: %~dp0media_scripts_setup.exe
    echo.
    echo  Run Setup.bat - it will auto-detect and use media_scripts_setup.exe.
    echo  Or run directly:
    echo    media_scripts_setup.exe
    echo    media_scripts_setup.exe --skip-handbrake --skip-staxrip
    echo    media_scripts_setup.exe --force
    echo ================================================================
) else (
    echo.
    echo ERROR: dist\media_scripts_setup.exe not found after build.
    pause
    exit /b 1
)

REM --- Clean up build/ dist/ spec (keep exe only) ---
if exist "build"                    rmdir /s /q "build"
if exist "dist"                     rmdir /s /q "dist"
if exist "media_scripts_setup.spec" del /q "media_scripts_setup.spec"

echo.
pause
