#Requires -Version 5.1
<#
.SYNOPSIS
    Media Scripts 2026 - First-time setup and tool installer.
.DESCRIPTION
    Creates the full directory structure, writes all bat files, and downloads
    all required tools. Safe to re-run: existing files are skipped unless
    -Force is specified.
.PARAMETER Force
    Re-download tools and overwrite bat files even if they already exist.
.PARAMETER DirectoriesOnly
    Only create the directory structure, skip bat files and downloads.
.PARAMETER SkipBatFiles
    Skip writing bat files (directories and tool downloads still run).
.PARAMETER SkipHandBrake
    Skip HandBrake download (large ~65 MB file).
.PARAMETER SkipStaxRip
    Skip StaxRip download.
#>
param(
    [switch]$Force,
    [switch]$DirectoriesOnly,
    [switch]$SkipBatFiles,
    [switch]$SkipHandBrake,
    [switch]$SkipStaxRip
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

Add-Type -AssemblyName System.IO.Compression.FileSystem

# ----------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------
function Write-Step  { param($msg) Write-Host "  $msg" -ForegroundColor Cyan }
function Write-OK    { param($msg) Write-Host "  [OK]  $msg" -ForegroundColor Green }
function Write-Skip  { param($msg) Write-Host "  [--]  $msg" -ForegroundColor Gray }
function Write-Fail  { param($msg) Write-Host "  [ERR] $msg" -ForegroundColor Red }
function Write-Warn  { param($msg) Write-Host "  [!]   $msg" -ForegroundColor Yellow }

function Get-GitHubLatest {
    param([string]$Repo)
    return Invoke-RestMethod "https://api.github.com/repos/$Repo/releases/latest" -UseBasicParsing
}

function Download-File {
    param([string]$Url, [string]$Dest, [string]$Label)
    Write-Step "Downloading $Label..."
    Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing
}

function Extract-Zip {
    param([string]$Zip, [string]$Dest)
    if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
    [System.IO.Compression.ZipFile]::ExtractToDirectory(
        (Resolve-Path $Zip).Path,
        (New-Item $Dest -ItemType Directory -Force).FullName
    )
}

function Get-7zTool {
    $7zr = Join-Path $root "core\7zr.exe"
    if (-not (Test-Path $7zr)) {
        Write-Step "Downloading 7zr.exe (~400 KB)..."
        Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7zr.exe' -OutFile $7zr -UseBasicParsing
    }
    return $7zr
}

function Extract-7z {
    param([string]$Archive, [string]$Dest)
    $7z = Get-7zTool
    if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
    New-Item $Dest -ItemType Directory -Force | Out-Null
    $outFlag = '-o' + (Resolve-Path $Dest).Path
    & $7z x $Archive $outFlag -y | Out-Null
}

function Write-BatFile {
    param([string]$RelPath, [string]$Content)
    $path = Join-Path $root $RelPath
    if ((Test-Path $path) -and -not $Force) {
        Write-Skip $RelPath
    } else {
        Set-Content -Path $path -Value $Content -Encoding ASCII
        Write-OK $RelPath
    }
}

# ----------------------------------------------------------------
# 1. Directory structure
# ----------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor White
Write-Host "  Media Scripts 2026 - Setup" -ForegroundColor White
Write-Host "  Root: $root" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor White
Write-Host ""
Write-Host "[1/4] Creating directory structure..." -ForegroundColor Yellow

$dirs = @(
    "Input",
    "Input\done",
    "Output",
    "Output\Music",
    "Output\Comics",
    "logs",
    "core\ffmpeg",
    "core\mpv",
    "core\staxrip"
)

foreach ($d in $dirs) {
    $path = Join-Path $root $d
    if (Test-Path $path) { Write-Skip $d }
    else { New-Item -Path $path -ItemType Directory -Force | Out-Null; Write-OK "Created: $d" }
}

if ($DirectoriesOnly) {
    Write-Host ""
    Write-Host "  Directory setup complete (-DirectoriesOnly)." -ForegroundColor Green
    exit 0
}

# ----------------------------------------------------------------
# 2. Create bat files
# ----------------------------------------------------------------
Write-Host ""
Write-Host "[2/4] Writing bat files..." -ForegroundColor Yellow

if ($SkipBatFiles) {
    Write-Skip "Bat file creation skipped (-SkipBatFiles)"
} else {

Write-BatFile "Download-Comics.bat" @'
@echo off
@title Comics Downloader
set working_dir=%cd%
echo Visit https://readallcomics.com and Copy the URL of the comic you want.
echo =======================================================================================
set /p userInput="Enter Comic URL: "
cls
set /p userFormat="Enter prefered download format (cbr , epub , pdf): "
cls
echo You entered: %userInput% %userFormat%
.\core\comics-dl.exe -url=%userInput% -format=%userFormat% -output=%cd%
cls
echo ==================================================================================
echo Download finished , press enter to quit
pause
cls
'@

Write-BatFile "Media_Rename_Only.bat" @'
.\core\python-3.12.3\Scripts\mnamer.exe --movie-directory="{name} ({year})" --episode-directory="{series}" .
'@

Write-BatFile "Setup_Path_Variables_If_Error.bat" @'
.\core\scripts\make_winpython_fix.bat
'@

Write-BatFile "Spotify-Downloader.bat" @'
@echo off
@title Spotify Downloader
echo Spotify-Downloader
echo =======================================================================================
set /p userInput="Enter Spotify Song or Playlist URL: "
cls
echo You entered: %userInput%
.\core\spotdl.exe %userInput%
cls
echo ==================================================================================
echo Download finished , press enter to quit
pause
'@

Write-BatFile "Stream-DL.bat" @'
@echo off
@title StreamDL Download
echo If you recieve a http 401 message browse the url in firefox to create a session cookie.
echo =======================================================================================
set /p userInput="Enter Streaming Video or Playlist URL: "
cls
echo You entered: %userInput%
.\core\youtube-dl.exe %userInput% --all-subs --cookies-from-browser firefox
cls
echo ==================================================================================
echo Download finished , press enter to attempt media rename based on file name or close window to quit
pause
cls
.\core\python-3.12.3\Scripts\mnamer.exe --movie-directory="{name} ({year})" --episode-directory="{series}" .
'@

Write-BatFile "Standard x264 Base Encode (HandBrake).bat" @'
@echo off
@title Encode - Standard x264 Base
setlocal
cd /d "%~dp0"

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  Standard x264 Base Encode - MP4 Output
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================
echo.
echo [%LOGSTAMP%] Standard x264 Base Encode started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in (".\Input\*.mp4" ".\Input\*.mkv" ".\Input\*.avi" ".\Input\*.mov" ".\Input\*.m4v" ".\Input\*.wmv" ".\Input\*.flv" ".\Input\*.webm" ".\Input\*.mpeg" ".\Input\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        .\core\HandBrakeCLI.exe --preset-import-file .\core\presets.json -Z "standard" ^
            -s "1,2,3,4,5,6" -i "%%i" -o ".\Output\%%~ni.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni.mp4 >> "%LOGFILE%"
            move "%%i" ".\Input\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
.\core\HandBrakeCLI.exe --preset-import-file .\core\presets.json -Z "standard" ^
    -s "1,2,3,4,5,6" -i "%INFILE%" -o ".\Output\%INNAME%.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

.\core\rename.bat
echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Standard x264 Base Encode (HandBrake MKV).bat" @'
@echo off
@title Encode - Standard x264 Base (HandBrake MKV)
setlocal
cd /d "%~dp0"

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  Standard x264 Base Encode - MKV Output
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================
echo.
echo [%LOGSTAMP%] Standard x264 Base Encode (MKV) started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in (".\Input\*.mp4" ".\Input\*.mkv" ".\Input\*.avi" ".\Input\*.mov" ".\Input\*.m4v" ".\Input\*.wmv" ".\Input\*.flv" ".\Input\*.webm" ".\Input\*.mpeg" ".\Input\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        .\core\HandBrakeCLI.exe --preset-import-file .\core\presets.json -Z "standard" ^
            --format av_mkv -s "1,2,3,4,5,6" -i "%%i" -o ".\Output\%%~ni.mkv"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni.mkv >> "%LOGFILE%"
            move "%%i" ".\Input\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni.mkv
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
.\core\HandBrakeCLI.exe --preset-import-file .\core\presets.json -Z "standard" ^
    --format av_mkv -s "1,2,3,4,5,6" -i "%INFILE%" -o ".\Output\%INNAME%.mkv"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%.mkv >> "%LOGFILE%"
    echo Done: %INNAME%.mkv
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

.\core\rename.bat
echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 480p Downscale (HandBrake).bat" @'
@echo off
@title Encode - 480p Downscale (HandBrake)
setlocal
cd /d "%~dp0"
set PRESET_FILE=.\core\480p.json
set PRESET_NAME=480pDownConvert
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  480p Downscale Encoder - x264, CRF 22, AAC 160k
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================
echo.
echo [%LOGSTAMP%] 480p Downscale (HandBrake) started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        .\core\HandBrakeCLI.exe --preset-import-file "%PRESET_FILE%" -Z "%PRESET_NAME%" ^
            -s "1,2,3,4,5,6" -i "%%i" -o "%OUTPUT_DIR%\%%~ni_480p.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_480p.mp4 >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_480p.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
.\core\HandBrakeCLI.exe --preset-import-file "%PRESET_FILE%" -Z "%PRESET_NAME%" ^
    -s "1,2,3,4,5,6" -i "%INFILE%" -o "%OUTPUT_DIR%\%INNAME%_480p.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_480p.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%_480p.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

.\core\rename.bat
echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 2160p Upscale (HandBrake).bat" @'
@echo off
@title Encode - 2160p Upscale (HandBrake)
setlocal
cd /d "%~dp0"
set PRESET_FILE=.\core\2160p.json
set PRESET_NAME=2160pUpscale
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  2160p Upscale Encoder - x265 HEVC, CRF 20, AAC 192k
echo  NOTE: Upscaling adds detail sharpening but cannot add detail
echo        that was not in the original source.
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================
echo.
echo [%LOGSTAMP%] 2160p Upscale (HandBrake) started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        .\core\HandBrakeCLI.exe --preset-import-file "%PRESET_FILE%" -Z "%PRESET_NAME%" ^
            -s "1,2,3,4,5,6" -i "%%i" -o "%OUTPUT_DIR%\%%~ni_2160p.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_2160p.mp4 >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_2160p.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
.\core\HandBrakeCLI.exe --preset-import-file "%PRESET_FILE%" -Z "%PRESET_NAME%" ^
    -s "1,2,3,4,5,6" -i "%INFILE%" -o "%OUTPUT_DIR%\%INNAME%_2160p.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_2160p.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%_2160p.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

.\core\rename.bat
echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 480p Downscale (FFmpeg).bat" @'
@echo off
@title Encode - 480p Downscale (FFmpeg)
setlocal
cd /d "%~dp0"
set FFMPEG=.\core\ffmpeg\ffmpeg.exe
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  480p Downscale - FFmpeg x264, CRF 22, scale to 854x480
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================

if not exist "%FFMPEG%" (
    echo ERROR: FFmpeg not found at %FFMPEG%
    echo Run Update-Tools.bat to download FFmpeg.
    pause
    exit /b 1
)

echo.
echo [%LOGSTAMP%] 480p Downscale (FFmpeg) started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        "%FFMPEG%" -hide_banner -loglevel warning -stats ^
            -i "%%i" ^
            -vf "scale=854:480:flags=lanczos,setsar=1" ^
            -c:v libx264 -crf 22 -preset medium -profile:v main ^
            -c:a aac -b:a 160k -ac 2 ^
            -c:s copy ^
            -movflags +faststart ^
            "%OUTPUT_DIR%\%%~ni_480p.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_480p.mp4 >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_480p.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
"%FFMPEG%" -hide_banner -loglevel warning -stats ^
    -i "%INFILE%" ^
    -vf "scale=854:480:flags=lanczos,setsar=1" ^
    -c:v libx264 -crf 22 -preset medium -profile:v main ^
    -c:a aac -b:a 160k -ac 2 ^
    -c:s copy ^
    -movflags +faststart ^
    "%OUTPUT_DIR%\%INNAME%_480p.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_480p.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%_480p.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 2160p Upscale (FFmpeg).bat" @'
@echo off
@title Encode - 2160p Upscale (FFmpeg)
setlocal
cd /d "%~dp0"
set FFMPEG=.\core\ffmpeg\ffmpeg.exe
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  2160p Upscale - FFmpeg x265 HEVC, CRF 20, scale to 3840x2160
echo  Uses Lanczos filter for high-quality upscaling.
echo  NOTE: x265 encoding is slow - this will take time.
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================

if not exist "%FFMPEG%" (
    echo ERROR: FFmpeg not found at %FFMPEG%
    echo Run Update-Tools.bat to download FFmpeg.
    pause
    exit /b 1
)

echo.
echo [%LOGSTAMP%] 2160p Upscale (FFmpeg) started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        "%FFMPEG%" -hide_banner -loglevel warning -stats ^
            -i "%%i" ^
            -vf "scale=3840:2160:flags=lanczos,setsar=1" ^
            -c:v libx265 -crf 20 -preset slow ^
            -tag:v hvc1 ^
            -c:a aac -b:a 192k -ac 2 ^
            -c:s copy ^
            -movflags +faststart ^
            "%OUTPUT_DIR%\%%~ni_2160p.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_2160p.mp4 >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_2160p.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
"%FFMPEG%" -hide_banner -loglevel warning -stats ^
    -i "%INFILE%" ^
    -vf "scale=3840:2160:flags=lanczos,setsar=1" ^
    -c:v libx265 -crf 20 -preset slow ^
    -tag:v hvc1 ^
    -c:a aac -b:a 192k -ac 2 ^
    -c:s copy ^
    -movflags +faststart ^
    "%OUTPUT_DIR%\%INNAME%_2160p.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_2160p.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%_2160p.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 480p Downscale (FFmpeg MKV).bat" @'
@echo off
@title Encode - 480p Downscale (FFmpeg MKV)
setlocal
cd /d "%~dp0"
set FFMPEG=.\core\ffmpeg\ffmpeg.exe
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  480p Downscale - FFmpeg x264, CRF 22, MKV output
echo  Subtitles: all tracks copied (ASS/SRT/PGS supported in MKV)
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================

if not exist "%FFMPEG%" (
    echo ERROR: FFmpeg not found at %FFMPEG%
    echo Run Update-Tools.bat to download FFmpeg.
    pause
    exit /b 1
)

echo.
echo [%LOGSTAMP%] 480p Downscale MKV started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        "%FFMPEG%" -hide_banner -loglevel warning -stats ^
            -i "%%i" ^
            -vf "scale=854:480:flags=lanczos,setsar=1" ^
            -c:v libx264 -crf 22 -preset medium -profile:v main ^
            -c:a aac -b:a 160k -ac 2 ^
            -c:s copy ^
            "%OUTPUT_DIR%\%%~ni_480p.mkv"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_480p.mkv >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_480p.mkv
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
"%FFMPEG%" -hide_banner -loglevel warning -stats ^
    -i "%INFILE%" ^
    -vf "scale=854:480:flags=lanczos,setsar=1" ^
    -c:v libx264 -crf 22 -preset medium -profile:v main ^
    -c:a aac -b:a 160k -ac 2 ^
    -c:s copy ^
    "%OUTPUT_DIR%\%INNAME%_480p.mkv"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_480p.mkv >> "%LOGFILE%"
    echo Done: %INNAME%_480p.mkv
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "Encode 2160p Upscale (FFmpeg MKV).bat" @'
@echo off
@title Encode - 2160p Upscale (FFmpeg MKV)
setlocal
cd /d "%~dp0"
set FFMPEG=.\core\ffmpeg\ffmpeg.exe
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  2160p Upscale - FFmpeg x265 HEVC, CRF 20, MKV output
echo  Subtitles: all tracks copied (ASS/SRT/PGS/VOBSUB supported)
echo  NOTE: x265 encoding is slow - this will take time.
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================

if not exist "%FFMPEG%" (
    echo ERROR: FFmpeg not found at %FFMPEG%
    echo Run Update-Tools.bat to download FFmpeg.
    pause
    exit /b 1
)

echo.
echo [%LOGSTAMP%] 2160p Upscale MKV started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Encoding: %%~nxi
        echo [%LOGSTAMP%] Encoding: %%~nxi >> "%LOGFILE%"
        "%FFMPEG%" -hide_banner -loglevel warning -stats ^
            -i "%%i" ^
            -vf "scale=3840:2160:flags=lanczos,setsar=1" ^
            -c:v libx265 -crf 20 -preset slow ^
            -c:a aac -b:a 192k -ac 2 ^
            -c:s copy ^
            "%OUTPUT_DIR%\%%~ni_2160p.mkv"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_2160p.mkv >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_2160p.mkv
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Encode failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Encoding: %~nx1
echo [%LOGSTAMP%] Encoding: %~nx1 >> "%LOGFILE%"
"%FFMPEG%" -hide_banner -loglevel warning -stats ^
    -i "%INFILE%" ^
    -vf "scale=3840:2160:flags=lanczos,setsar=1" ^
    -c:v libx265 -crf 20 -preset slow ^
    -c:a aac -b:a 192k -ac 2 ^
    -c:s copy ^
    "%OUTPUT_DIR%\%INNAME%_2160p.mkv"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_2160p.mkv >> "%LOGFILE%"
    echo Done: %INNAME%_2160p.mkv
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Encode failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "SBS to Anaglyph 3D.bat" @'
@echo off
@title Encode - SBS to Anaglyph 3D
setlocal
cd /d "%~dp0"
set FFMPEG=.\core\ffmpeg\ffmpeg.exe
set INPUT_DIR=.\Input
set OUTPUT_DIR=.\Output

for /f "tokens=*" %%d in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set LOGSTAMP=%%d
set LOGFILE=.\logs\encode_%LOGSTAMP%.log

echo ================================================================
echo  SBS to Anaglyph 3D - Half Side-by-Side to Red/Cyan Dubois
echo  Input:  Half SBS (left eye on left)
echo  Output: Anaglyph Red/Cyan (Dubois) - x264, CRF 20, MP4
echo  Log: logs\encode_%LOGSTAMP%.log
echo ================================================================

if not exist "%FFMPEG%" (
    echo ERROR: FFmpeg not found at %FFMPEG%
    echo Run Update-Tools.bat to download FFmpeg.
    pause
    exit /b 1
)

echo.
echo [%LOGSTAMP%] SBS to Anaglyph 3D started >> "%LOGFILE%"

set found=0

if not "%~1"=="" goto :drag_mode

for %%i in ("%INPUT_DIR%\*.mp4" "%INPUT_DIR%\*.mkv" "%INPUT_DIR%\*.avi" "%INPUT_DIR%\*.mov" "%INPUT_DIR%\*.m4v" "%INPUT_DIR%\*.wmv" "%INPUT_DIR%\*.flv" "%INPUT_DIR%\*.webm" "%INPUT_DIR%\*.mpeg" "%INPUT_DIR%\*.mpg") do (
    if exist "%%i" (
        set found=1
        echo Converting: %%~nxi
        echo [%LOGSTAMP%] Converting: %%~nxi >> "%LOGFILE%"
        "%FFMPEG%" -hide_banner -loglevel warning -stats ^
            -i "%%i" ^
            -vf "stereo3d=sbs2l:arcd" ^
            -c:v libx264 -crf 20 -preset fast ^
            -c:a copy ^
            -movflags +faststart ^
            "%OUTPUT_DIR%\%%~ni_Anaglyph.mp4"
        if not errorlevel 1 (
            echo [%LOGSTAMP%] OK: %%~ni_Anaglyph.mp4 >> "%LOGFILE%"
            move "%%i" "%INPUT_DIR%\done\" >nul 2>&1
            if errorlevel 1 del "%%i"
            echo Done: %%~ni_Anaglyph.mp4
        ) else (
            echo [%LOGSTAMP%] FAILED: %%~nxi >> "%LOGFILE%"
            echo WARNING: Conversion failed for %%~nxi - file kept in Input
        )
        echo.
    )
)
goto :all_done

:drag_mode
if "%~1"=="" goto :all_done
set "INFILE=%~1"
set "INNAME=%~n1"
set found=1
echo Converting: %~nx1
echo [%LOGSTAMP%] Converting: %~nx1 >> "%LOGFILE%"
"%FFMPEG%" -hide_banner -loglevel warning -stats ^
    -i "%INFILE%" ^
    -vf "stereo3d=sbs2l:arcd" ^
    -c:v libx264 -crf 20 -preset fast ^
    -c:a copy ^
    -movflags +faststart ^
    "%OUTPUT_DIR%\%INNAME%_Anaglyph.mp4"
if not errorlevel 1 (
    echo [%LOGSTAMP%] OK: %INNAME%_Anaglyph.mp4 >> "%LOGFILE%"
    echo Done: %INNAME%_Anaglyph.mp4
) else (
    echo [%LOGSTAMP%] FAILED: %~nx1 >> "%LOGFILE%"
    echo WARNING: Conversion failed for %~nx1
)
echo.
shift
goto :drag_mode

:all_done
if "%found%"=="0" (
    echo No video files found.
    echo Drag a SBS 3D file onto this bat, or place files in .\Input\
    echo Supported: mp4 mkv avi mov m4v wmv flv webm mpeg mpg
    echo [%LOGSTAMP%] No input files found >> "%LOGFILE%"
)

echo [%LOGSTAMP%] Session complete >> "%LOGFILE%"
echo.
echo ================================================================
echo  All done. Press any key to exit.
pause >nul
'@

Write-BatFile "FFmpeg Prompt.bat" @'
@echo off
@title FFmpeg Command Prompt
setlocal

set FFMPEG_DIR=%~dp0core\ffmpeg
set MPV_DIR=%~dp0core\mpv

if not exist "%FFMPEG_DIR%\ffmpeg.exe" (
    echo WARNING: FFmpeg not found at %FFMPEG_DIR%
    echo Run Update-Tools.bat to download FFmpeg first.
    echo.
)

REM Add ffmpeg and mpv to PATH for this session
set PATH=%FFMPEG_DIR%;%MPV_DIR%;%PATH%

echo ================================================================
echo  FFmpeg / MPV Command Prompt
echo ================================================================
echo  ffmpeg.exe and mpv.exe are available in this shell.
echo  Input files:  %~dp0Input\
echo  Output files: %~dp0Output\
echo.
echo  Quick reference:
echo    ffmpeg -i input.mp4 -vf scale=854:480 -c:v libx264 -crf 22 out.mp4
echo    ffmpeg -i input.mp4 -vf scale=3840:2160:flags=lanczos -c:v libx265 -crf 20 out.mp4
echo    ffmpeg -i input.mp4 -ss 00:01:00 -t 00:00:30 clip.mp4
echo    ffprobe -v quiet -print_format json -show_format -show_streams input.mp4
echo    mpv input.mp4
echo ================================================================
echo.

cd /d "%~dp0"
cmd /k "echo Working directory: %~dp0 && echo."
endlocal
'@

Write-BatFile "Open MPV.bat" @'
@echo off
@title MPV Player
setlocal
cd /d "%~dp0"
set MPV=.\core\mpv\mpv.exe

if not exist "%MPV%" (
    echo ERROR: MPV not found at %MPV%
    echo Run Update-Tools.bat to download MPV.
    echo Or manually place mpv.exe in .\core\mpv\
    pause
    exit /b 1
)

if "%~1"=="" (
    REM No file argument - open file picker via PowerShell
    echo Launching MPV file picker...
    for /f "delims=" %%f in ('powershell -NoProfile -STA -Command ^
        "Add-Type -AssemblyName System.Windows.Forms;" ^
        "[System.Windows.Forms.Application]::EnableVisualStyles();" ^
        "$f = New-Object System.Windows.Forms.OpenFileDialog;" ^
        "$f.Title = 'Select video to play with MPV';" ^
        "$f.Filter = 'Video Files|*.mp4;*.mkv;*.avi;*.mov;*.m4v;*.wmv;*.flv;*.webm;*.mpeg;*.mpg;*.ts|All Files|*.*';" ^
        "if ($f.ShowDialog() -eq 'OK') { $f.FileName }"') do (
        "%MPV%" --really-quiet "%%f"
        goto :end
    )
    echo No file selected.
) else (
    "%MPV%" --really-quiet %*
)

:end
endlocal
'@

Write-BatFile "Open StaxRip.bat" @'
@echo off
@title StaxRip
setlocal

set STAXRIP=.\core\staxrip\StaxRip.exe

if exist "%STAXRIP%" (
    echo Launching StaxRip...
    start "" "%STAXRIP%"
    goto :end
)

echo StaxRip not found at .\core\staxrip\StaxRip.exe
echo.
echo Run Update-Tools.bat and select StaxRip to download automatically,
echo or download manually from: https://github.com/staxrip/staxrip/releases
echo.
echo After downloading, extract StaxRip.exe to:
echo   %~dp0core\staxrip\StaxRip.exe
echo.

:end
pause
endlocal
'@

# core helper bats
Write-BatFile "core\rename.bat" @'
.\core\python-3.12.3\Scripts\mnamer.exe --batch --no-overwrite --movie-directory="{name} ({year})" --episode-directory="{series}" output
'@

Write-BatFile "core\mnamer.bat" @'
.\core\python-3.12.3\Scripts\mnamer.exe --movie-directory="{name} ({year})" --episode-directory="{series}" .
'@

# Update-Tools.bat - written last as it is the largest file
Write-BatFile "Update-Tools.bat" @'
@echo off
@title Media Tools Updater
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo ================================================================
echo  Media Tools Version Checker ^& Updater
echo ================================================================
echo.

REM ---- Show current installed versions ----
echo [Installed Versions]
echo.
call :show_versions
echo.

echo ================================================================
echo  What would you like to update?
echo  Enter one or multiple numbers separated by spaces or commas.
echo  Examples:  1        (just yt-dlp)
echo             1 3 4    (yt-dlp, FFmpeg, MPV)
echo             1,3,4    (same with commas)
echo ================================================================
echo   [1] yt-dlp  ^(saved as youtube-dl.exe for legacy compat^)
echo   [2] spotdl
echo   [3] FFmpeg  ^(ffmpeg.exe + ffprobe.exe^)
echo   [4] MPV
echo   [5] HandBrakeCLI
echo   [6] StaxRip
echo   [7] comics-dl
echo   [8] 7-Zip  ^(7zr.exe - used for .7z extraction^)
echo   [9] Rickinator  ^(link-list download manager^)
echo  [10] All of the above
echo   [0] Exit
echo.
set /p choice="Enter choice(s): "

REM Normalise: replace commas with spaces so "for" can tokenise either format
set "choice=%choice:,= %"

if "%choice%"=="0" goto :end

set "any_valid=0"
for %%c in (%choice%) do (
    if "%%c"=="0" goto :end
    if "%%c"=="1"  ( call :update_ytdlp      & set "any_valid=1" )
    if "%%c"=="2"  ( call :update_spotdl     & set "any_valid=1" )
    if "%%c"=="3"  ( call :update_ffmpeg     & set "any_valid=1" )
    if "%%c"=="4"  ( call :update_mpv        & set "any_valid=1" )
    if "%%c"=="5"  ( call :update_handbrake  & set "any_valid=1" )
    if "%%c"=="6"  ( call :update_staxrip    & set "any_valid=1" )
    if "%%c"=="7"  ( call :update_comicsdl   & set "any_valid=1" )
    if "%%c"=="8"  ( call :update_7zr        & set "any_valid=1" )
    if "%%c"=="9"  ( call :update_rickinator & set "any_valid=1" )
    if "%%c"=="10" (
        call :update_ytdlp
        call :update_spotdl
        call :update_ffmpeg
        call :update_mpv
        call :update_handbrake
        call :update_staxrip
        call :update_comicsdl
        call :update_7zr
        call :update_rickinator
        set "any_valid=1"
    )
)
if "!any_valid!"=="1" goto :summary
echo Invalid choice.
goto :end

REM ================================================================
:show_versions
set "V=NOT INSTALLED"
if exist ".\core\youtube-dl.exe" for /f "tokens=*" %%v in ('.\core\youtube-dl.exe --version 2^>^&1') do set "V=%%v"
echo   yt-dlp ^(youtube-dl.exe^): !V!

set "V=NOT INSTALLED"
if exist ".\core\spotdl.exe" for /f "tokens=*" %%v in ('.\core\spotdl.exe --version 2^>^&1') do set "V=%%v"
echo   spotdl: !V!

set "V=NOT INSTALLED"
if exist ".\core\HandBrakeCLI.exe" call :_ver_handbrake
echo   HandBrakeCLI: !V!

set "V=NOT INSTALLED"
if exist ".\core\ffmpeg\ffmpeg.exe" call :_ver_ffmpeg
echo   FFmpeg: !V!

set "V=NOT INSTALLED"
if exist ".\core\mpv\mpv.exe" call :_ver_mpv
echo   MPV: !V!

if exist ".\core\staxrip\StaxRip.exe" ( echo   StaxRip: INSTALLED ) else echo   StaxRip: NOT INSTALLED

set "V=NOT INSTALLED"
if exist ".\core\comics-dl.exe" for /f "tokens=*" %%v in ('.\core\comics-dl.exe --version 2^>^&1') do set "V=%%v"
echo   comics-dl: !V!

set "V=NOT INSTALLED"
if exist ".\core\7zr.exe" call :_ver_7zr
echo   7-Zip (7zr.exe): !V!

if exist ".\Rickinator.exe" ( echo   Rickinator: INSTALLED ) else echo   Rickinator: NOT INSTALLED
exit /b

:_ver_handbrake
for /f "tokens=1,2,3" %%a in ('.\core\HandBrakeCLI.exe --version 2^>^&1 ^| findstr /i "HandBrake"') do set "V=%%a %%b %%c"
exit /b

:_ver_ffmpeg
for /f "tokens=3" %%v in ('.\core\ffmpeg\ffmpeg.exe -version 2^>^&1 ^| findstr "ffmpeg version"') do set "V=%%v"
exit /b

:_ver_mpv
for /f "tokens=1,2" %%a in ('.\core\mpv\mpv.exe --version 2^>^&1 ^| findstr /i "mpv v"') do set "V=%%a %%b"
exit /b

:_ver_7zr
for /f "tokens=3" %%v in ('.\core\7zr.exe 2^>^&1 ^| findstr "7-Zip"') do set "V=%%v"
exit /b

REM ================================================================
:update_ytdlp
echo.
echo [yt-dlp] Downloading latest release...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $url = 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe';" ^
    "  $dest = '.\core\youtube-dl.exe';" ^
    "  Write-Host ('  Source: ' + $url);" ^
    "  Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing;" ^
    "  Write-Host '  Saved as youtube-dl.exe (legacy compatibility name)';" ^
    "  $ver = & $dest '--version' 2>&1 | Select-Object -First 1;" ^
    "  Write-Host ('  Installed version: ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_spotdl
echo.
echo [spotdl] Downloading latest release...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/spotDL/spotify-downloader/releases/latest' -UseBasicParsing;" ^
    "  $asset = $api.assets | Where-Object { $_.name -match '^spotdl.*windows.*\.exe$' } | Select-Object -First 1;" ^
    "  if (-not $asset) { $asset = $api.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1; }" ^
    "  if (-not $asset) { Write-Host '  ERROR: Could not find Windows exe in release assets'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name);" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile '.\core\spotdl.exe' -UseBasicParsing;" ^
    "  $ver = & '.\core\spotdl.exe' '--version' 2>&1 | Select-Object -First 1;" ^
    "  Write-Host ('  Installed version: ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_ffmpeg
echo.
echo [FFmpeg] Downloading latest essentials build...
if not exist ".\core\ffmpeg" mkdir ".\core\ffmpeg"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  Add-Type -AssemblyName System.IO.Compression.FileSystem;" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/GyanD/codexffmpeg/releases/latest' -UseBasicParsing;" ^
    "  Write-Host ('  Latest release: ' + $api.tag_name);" ^
    "  $asset = $api.assets | Where-Object { $_.name -match 'essentials_build.*\.zip$' } | Select-Object -First 1;" ^
    "  if (-not $asset) { Write-Host '  ERROR: No essentials zip found in release.'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name + ' (' + [math]::Round($asset.size/1MB,1) + ' MB)');" ^
    "  $zip = '.\core\ffmpeg\_dl.zip';" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -UseBasicParsing;" ^
    "  $tmp = '.\core\ffmpeg\_tmp';" ^
    "  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force; }" ^
    "  [System.IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path $zip).Path, (New-Item $tmp -ItemType Directory -Force).FullName);" ^
    "  $exes = Get-ChildItem $tmp -Recurse -Include 'ffmpeg.exe','ffprobe.exe';" ^
    "  foreach ($e in $exes) { Copy-Item $e.FullName '.\core\ffmpeg\' -Force; Write-Host ('  Installed: ' + $e.Name); }" ^
    "  Remove-Item $zip -Force; Remove-Item $tmp -Recurse -Force;" ^
    "  $ver = & '.\core\ffmpeg\ffmpeg.exe' '-version' 2>&1 | Select-String 'ffmpeg version' | Select-Object -First 1;" ^
    "  Write-Host ('  ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_mpv
echo.
echo [MPV] Downloading latest Windows build...
if not exist ".\core\mpv" mkdir ".\core\mpv"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/zhongfly/mpv-winbuild/releases/latest' -UseBasicParsing;" ^
    "  Write-Host ('  Latest release: ' + $api.tag_name);" ^
    "  $asset = $api.assets | Where-Object { $_.name -match 'mpv-x86_64.*\.7z$' } | Select-Object -First 1;" ^
    "  if (-not $asset) { Write-Host '  ERROR: No x64 .7z found. Check https://github.com/zhongfly/mpv-winbuild/releases manually.'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name + ' (' + [math]::Round($asset.size/1MB,1) + ' MB)');" ^
    "  $archive = '.\core\mpv\_dl.7z';" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $archive -UseBasicParsing;" ^
    "  if (-not (Test-Path '.\core\7zr.exe')) {" ^
    "    Write-Host '  7zr.exe not found - downloading...';" ^
    "    Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7zr.exe' -OutFile '.\core\7zr.exe' -UseBasicParsing;" ^
    "  }" ^
    "  Write-Host '  Extracting...';" ^
    "  $tmp = '.\core\mpv\_tmp';" ^
    "  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force; }" ^
    "  New-Item $tmp -ItemType Directory -Force | Out-Null;" ^
    "  $outFlag = '-o' + (Resolve-Path $tmp).Path;" ^
    "  & '.\core\7zr.exe' x $archive $outFlag -y | Out-Null;" ^
    "  $exe = Get-ChildItem $tmp -Recurse -Include 'mpv.exe' | Select-Object -First 1;" ^
    "  if ($exe) { Copy-Item $exe.FullName '.\core\mpv\' -Force; Write-Host ('  Installed: ' + $exe.Name); }" ^
    "  Remove-Item $archive -Force; Remove-Item $tmp -Recurse -Force;" ^
    "  Write-Host '  MPV updated successfully.';" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_handbrake
echo.
echo [HandBrakeCLI] Downloading latest release (~65 MB)...
set /p hb_confirm="This is a large download. Continue? (y/n): "
if /i not "%hb_confirm%"=="y" ( echo Skipped. & exit /b )
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/HandBrake/HandBrake/releases/latest' -UseBasicParsing;" ^
    "  Write-Host ('  Latest release: ' + $api.tag_name);" ^
    "  $asset = $api.assets | Where-Object { $_.name -match 'HandBrakeCLI.*x86_64.*\.zip$' } | Select-Object -First 1;" ^
    "  if (-not $asset) { Write-Host '  ERROR: Could not find HandBrakeCLI zip. Check https://handbrake.fr/downloads2.php'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name + ' (' + [math]::Round($asset.size/1MB,1) + ' MB)');" ^
    "  Add-Type -AssemblyName System.IO.Compression.FileSystem;" ^
    "  $zip = '.\core\_hb_dl.zip';" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -UseBasicParsing;" ^
    "  $tmp = '.\core\_hb_tmp';" ^
    "  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force; }" ^
    "  [System.IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path $zip).Path, (New-Item $tmp -ItemType Directory -Force).FullName);" ^
    "  $exe = Get-ChildItem $tmp -Recurse -Include 'HandBrakeCLI.exe' | Select-Object -First 1;" ^
    "  if ($exe) { Copy-Item $exe.FullName '.\core\HandBrakeCLI.exe' -Force; Write-Host '  HandBrakeCLI.exe updated.'; }" ^
    "  Remove-Item $zip -Force; Remove-Item $tmp -Recurse -Force;" ^
    "  $ver = & '.\core\HandBrakeCLI.exe' '--version' 2>&1 | Select-String 'HandBrake' | Select-Object -First 1;" ^
    "  Write-Host ('  ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_staxrip
echo.
echo [StaxRip] Downloading latest release...
if not exist ".\core\staxrip" mkdir ".\core\staxrip"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/staxrip/staxrip/releases/latest' -UseBasicParsing;" ^
    "  Write-Host ('  Latest release: ' + $api.tag_name);" ^
    "  $asset = $api.assets | Where-Object { $_.name -match 'StaxRip.*x64.*\.7z$' } | Select-Object -First 1;" ^
    "  if (-not $asset) { $asset = $api.assets | Where-Object { $_.name -match '\.7z$' } | Select-Object -First 1; }" ^
    "  if (-not $asset) { Write-Host '  ERROR: No .7z asset found. Check https://github.com/staxrip/staxrip/releases'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name + ' (' + [math]::Round($asset.size/1MB,1) + ' MB)');" ^
    "  $archive = '.\core\staxrip\_dl.7z';" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $archive -UseBasicParsing;" ^
    "  if (-not (Test-Path '.\core\7zr.exe')) {" ^
    "    Write-Host '  7zr.exe not found - downloading...';" ^
    "    Invoke-WebRequest -Uri 'https://www.7-zip.org/a/7zr.exe' -OutFile '.\core\7zr.exe' -UseBasicParsing;" ^
    "  }" ^
    "  Write-Host '  Extracting...';" ^
    "  $tmp = '.\core\staxrip\_tmp';" ^
    "  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force; }" ^
    "  New-Item $tmp -ItemType Directory -Force | Out-Null;" ^
    "  $outFlag = '-o' + (Resolve-Path $tmp).Path;" ^
    "  & '.\core\7zr.exe' x $archive $outFlag -y | Out-Null;" ^
    "  $sub = Get-ChildItem $tmp -Directory | Select-Object -First 1;" ^
    "  $src = if ($sub) { $sub.FullName } else { $tmp };" ^
    "  Get-ChildItem $src | ForEach-Object { Copy-Item $_.FullName '.\core\staxrip\' -Recurse -Force; }" ^
    "  Remove-Item $archive -Force; Remove-Item $tmp -Recurse -Force;" ^
    "  Write-Host '  StaxRip installed to .\core\staxrip\';" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_7zr
echo.
echo [7-Zip] Downloading latest 7zr.exe...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $url = 'https://www.7-zip.org/a/7zr.exe';" ^
    "  Write-Host ('  Source: ' + $url);" ^
    "  Invoke-WebRequest -Uri $url -OutFile '.\core\7zr.exe' -UseBasicParsing;" ^
    "  $ver = & '.\core\7zr.exe' 2>&1 | Select-String '7-Zip' | Select-Object -First 1;" ^
    "  Write-Host ('  Installed: ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_comicsdl
echo.
echo [comics-dl] Downloading latest release from Girbons/comics-downloader...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  $api = Invoke-RestMethod 'https://api.github.com/repos/Girbons/comics-downloader/releases/latest' -UseBasicParsing;" ^
    "  Write-Host ('  Latest release: ' + $api.tag_name);" ^
    "  $asset = $api.assets | Where-Object { $_.name -match 'windows.*amd64.*\.exe$' -or $_.name -match 'windows-amd64' } | Select-Object -First 1;" ^
    "  if (-not $asset) { $asset = $api.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1; }" ^
    "  if (-not $asset) { Write-Host '  ERROR: Could not find Windows exe. Check https://github.com/Girbons/comics-downloader/releases'; return; }" ^
    "  Write-Host ('  Downloading: ' + $asset.name);" ^
    "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile '.\core\comics-dl.exe' -UseBasicParsing;" ^
    "  Write-Host '  Saved as comics-dl.exe';" ^
    "  $ver = & '.\core\comics-dl.exe' '--version' 2>&1 | Select-Object -First 1;" ^
    "  Write-Host ('  Installed version: ' + $ver);" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:update_rickinator
echo.
echo [Rickinator] Downloading...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try {" ^
    "  Add-Type -AssemblyName System.IO.Compression.FileSystem;" ^
    "  $url = 'https://github.com/pasiegel/Rickinator/releases/download/downloader/Rickinator-1.0.zip';" ^
    "  Write-Host '  Downloading: Rickinator-1.0.zip';" ^
    "  $zip = '.\Rickinator-1.0.zip';" ^
    "  Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing;" ^
    "  Write-Host '  Extracting to root...';" ^
    "  $tmp = '.\Rickinator_tmp';" ^
    "  if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force; }" ^
    "  [System.IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path $zip).Path, (New-Item $tmp -ItemType Directory -Force).FullName);" ^
    "  $sub = Get-ChildItem $tmp -Directory | Select-Object -First 1;" ^
    "  $src = if ($sub) { $sub.FullName } else { $tmp };" ^
    "  Get-ChildItem $src | ForEach-Object { Copy-Item $_.FullName '.\' -Recurse -Force; }; Remove-Item $zip -Force; Remove-Item $tmp -Recurse -Force;" ^
    "  Write-Host '  Rickinator installed to root directory.';" ^
    "} catch { Write-Host ('  ERROR: ' + $_.Exception.Message) }"
exit /b

REM ================================================================
:summary
echo.
echo ================================================================
echo  Updated Versions
echo ================================================================
call :show_versions
echo.

:end
echo.
pause
endlocal
'@

} # end -SkipBatFiles

# ----------------------------------------------------------------
# 3. Download tools
# ----------------------------------------------------------------
Write-Host ""
Write-Host "[3/4] Downloading tools..." -ForegroundColor Yellow

$results = [ordered]@{}

# --- yt-dlp -> youtube-dl.exe ---
$dest = Join-Path $root "core\youtube-dl.exe"
if ((Test-Path $dest) -and -not $Force) {
    $ver = & $dest '--version' 2>&1 | Select-Object -First 1
    Write-Skip "yt-dlp already installed ($ver)"
    $results['yt-dlp'] = "skipped ($ver)"
} else {
    try {
        Download-File 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe' $dest "yt-dlp"
        $ver = & $dest '--version' 2>&1 | Select-Object -First 1
        Write-OK "yt-dlp installed as youtube-dl.exe ($ver)"
        $results['yt-dlp'] = "installed ($ver)"
    } catch { Write-Fail "yt-dlp: $($_.Exception.Message)"; $results['yt-dlp'] = "FAILED" }
}

# --- spotdl ---
$dest = Join-Path $root "core\spotdl.exe"
if ((Test-Path $dest) -and -not $Force) {
    $ver = & $dest '--version' 2>&1 | Select-Object -First 1
    Write-Skip "spotdl already installed ($ver)"
    $results['spotdl'] = "skipped ($ver)"
} else {
    try {
        $rel   = Get-GitHubLatest "spotDL/spotify-downloader"
        $asset = $rel.assets | Where-Object { $_.name -match 'windows.*\.exe$' -or $_.name -match '^spotdl.*\.exe$' } | Select-Object -First 1
        if (-not $asset) { throw "No Windows exe found in release assets" }
        Download-File $asset.browser_download_url $dest "spotdl $($rel.tag_name)"
        $ver = & $dest '--version' 2>&1 | Select-Object -First 1
        Write-OK "spotdl installed ($ver)"
        $results['spotdl'] = "installed ($ver)"
    } catch { Write-Fail "spotdl: $($_.Exception.Message)"; $results['spotdl'] = "FAILED" }
}

# --- 7zr.exe ---
$dest = Join-Path $root "core\7zr.exe"
if ((Test-Path $dest) -and -not $Force) {
    Write-Skip "7zr.exe already installed"
    $results['7-Zip (7zr)'] = "skipped (exists)"
} else {
    try {
        Download-File 'https://www.7-zip.org/a/7zr.exe' $dest "7zr.exe"
        $ver = (& $dest 2>&1 | Select-String '7-Zip' | Select-Object -First 1).ToString().Trim()
        Write-OK "7zr.exe installed ($ver)"
        $results['7-Zip (7zr)'] = "installed"
    } catch { Write-Fail "7zr.exe: $($_.Exception.Message)"; $results['7-Zip (7zr)'] = "FAILED" }
}

# --- pycore (portable Python + mnamer) ---
$dest = Join-Path $root "core\python-3.12.3"
if ((Test-Path $dest) -and -not $Force) {
    Write-Skip "pycore (Python + mnamer) already installed"
    $results['pycore'] = "skipped (exists)"
} else {
    try {
        $archive = Join-Path $root "core\_pycore.7z"
        Download-File 'https://github.com/pasiegel/Media-Scripts-2026/raw/refs/heads/main/resources/pycore.7z' $archive "pycore - Python + mnamer"
        $7z      = Get-7zTool
        Write-Step "Extracting pycore into core\..."
        $outFlag = '-o' + (Join-Path $root "core")
        & $7z x $archive $outFlag -y | Out-Null
        Remove-Item $archive -Force
        Write-OK "pycore installed (python-3.12.3 + mnamer)"
        $results['pycore'] = "installed"
    } catch {
        Write-Fail "pycore: $($_.Exception.Message)"
        $results['pycore'] = "FAILED"
    }
}

# --- Patch portable Python paths ---
# Must run after pycore is present. Re-run any time the folder is moved.
$fixBat = Join-Path $root "core\scripts\make_winpython_fix.bat"
if (Test-Path $fixBat) {
    Write-Step "Patching portable Python paths (make_winpython_fix)..."
    try {
        & cmd.exe /c "`"$fixBat`"" 2>&1 | Out-Null
        Write-OK "Python paths patched - mnamer ready"
        $results['Python path fix'] = "OK"
    } catch {
        Write-Fail "Python path fix failed: $($_.Exception.Message)"
        $results['Python path fix'] = "FAILED - run Setup_Path_Variables_If_Error.bat manually"
    }
} else {
    Write-Warn "core\scripts\make_winpython_fix.bat not found - pycore may not be extracted yet"
    Write-Warn "If mnamer fails, run Setup_Path_Variables_If_Error.bat"
    $results['Python path fix'] = "skipped (scripts not found)"
}

# --- FFmpeg ---
$dest = Join-Path $root "core\ffmpeg\ffmpeg.exe"
if ((Test-Path $dest) -and -not $Force) {
    Write-Skip "FFmpeg already installed"
    $results['FFmpeg'] = "skipped"
} else {
    try {
        $rel   = Get-GitHubLatest "GyanD/codexffmpeg"
        $asset = $rel.assets | Where-Object { $_.name -match 'essentials_build.*\.zip$' } | Select-Object -First 1
        if (-not $asset) { throw "No essentials zip found" }
        $zip = Join-Path $root "core\ffmpeg\_dl.zip"
        $tmp = Join-Path $root "core\ffmpeg\_tmp"
        Download-File $asset.browser_download_url $zip "FFmpeg $($rel.tag_name) ($([math]::Round($asset.size/1MB,1)) MB)"
        Extract-Zip $zip $tmp
        Get-ChildItem $tmp -Recurse -Include 'ffmpeg.exe','ffprobe.exe' | ForEach-Object { Copy-Item $_.FullName (Join-Path $root "core\ffmpeg\") -Force }
        Remove-Item $zip -Force; Remove-Item $tmp -Recurse -Force
        Write-OK "FFmpeg installed"
        $results['FFmpeg'] = "installed"
    } catch { Write-Fail "FFmpeg: $($_.Exception.Message)"; $results['FFmpeg'] = "FAILED" }
}

# --- MPV ---
$dest = Join-Path $root "core\mpv\mpv.exe"
if ((Test-Path $dest) -and -not $Force) {
    Write-Skip "MPV already installed"
    $results['MPV'] = "skipped"
} else {
    try {
        $rel   = Get-GitHubLatest "zhongfly/mpv-winbuild"
        $asset = $rel.assets | Where-Object { $_.name -match 'mpv-x86_64.*\.7z$' } | Select-Object -First 1
        if (-not $asset) { throw "No x64 .7z found in release" }
        $archive = Join-Path $root "core\mpv\_dl.7z"
        $tmp     = Join-Path $root "core\mpv\_tmp"
        Download-File $asset.browser_download_url $archive "MPV $($rel.tag_name) ($([math]::Round($asset.size/1MB,1)) MB)"
        Extract-7z $archive $tmp
        $exe = Get-ChildItem $tmp -Recurse -Include 'mpv.exe' | Select-Object -First 1
        if ($exe) { Copy-Item $exe.FullName (Join-Path $root "core\mpv\") -Force }
        Remove-Item $archive -Force; Remove-Item $tmp -Recurse -Force
        Write-OK "MPV installed"
        $results['MPV'] = "installed"
    } catch { Write-Fail "MPV: $($_.Exception.Message)"; $results['MPV'] = "FAILED" }
}

# --- HandBrakeCLI ---
$dest = Join-Path $root "core\HandBrakeCLI.exe"
if ($SkipHandBrake) {
    Write-Skip "HandBrake skipped (-SkipHandBrake)"
    $results['HandBrakeCLI'] = "skipped"
} elseif ((Test-Path $dest) -and -not $Force) {
    Write-Skip "HandBrakeCLI already installed"
    $results['HandBrakeCLI'] = "skipped (exists)"
} else {
    try {
        $rel   = Get-GitHubLatest "HandBrake/HandBrake"
        $asset = $rel.assets | Where-Object { $_.name -match 'HandBrakeCLI.*x86_64.*\.zip$' } | Select-Object -First 1
        if (-not $asset) { throw "No HandBrakeCLI zip found. Visit https://handbrake.fr/downloads2.php" }
        $zip = Join-Path $root "core\_hb_dl.zip"
        $tmp = Join-Path $root "core\_hb_tmp"
        Download-File $asset.browser_download_url $zip "HandBrakeCLI $($rel.tag_name) ($([math]::Round($asset.size/1MB,1)) MB)"
        Extract-Zip $zip $tmp
        $exe = Get-ChildItem $tmp -Recurse -Include 'HandBrakeCLI.exe' | Select-Object -First 1
        if ($exe) { Copy-Item $exe.FullName $dest -Force }
        Remove-Item $zip -Force; Remove-Item $tmp -Recurse -Force
        Write-OK "HandBrakeCLI installed"
        $results['HandBrakeCLI'] = "installed"
    } catch { Write-Fail "HandBrakeCLI: $($_.Exception.Message)"; $results['HandBrakeCLI'] = "FAILED" }
}

# --- StaxRip ---
$dest = Join-Path $root "core\staxrip\StaxRip.exe"
if ($SkipStaxRip) {
    Write-Skip "StaxRip skipped (-SkipStaxRip)"
    $results['StaxRip'] = "skipped"
} elseif ((Test-Path $dest) -and -not $Force) {
    Write-Skip "StaxRip already installed"
    $results['StaxRip'] = "skipped (exists)"
} else {
    try {
        $rel   = Get-GitHubLatest "staxrip/staxrip"
        $asset = $rel.assets | Where-Object { $_.name -match 'StaxRip.*x64.*\.7z$' } | Select-Object -First 1
        if (-not $asset) { $asset = $rel.assets | Where-Object { $_.name -match '\.7z$' } | Select-Object -First 1 }
        if (-not $asset) { throw "No .7z asset found. Check https://github.com/staxrip/staxrip/releases" }
        $archive = Join-Path $root "core\staxrip\_dl.7z"
        $tmp     = Join-Path $root "core\staxrip\_tmp"
        Download-File $asset.browser_download_url $archive "StaxRip $($rel.tag_name) ($([math]::Round($asset.size/1MB,1)) MB)"
        Extract-7z $archive $tmp
        $sub = Get-ChildItem $tmp -Directory | Select-Object -First 1
        $src = if ($sub) { $sub.FullName } else { $tmp }
        Get-ChildItem $src | ForEach-Object { Copy-Item $_.FullName (Join-Path $root "core\staxrip\") -Recurse -Force }
        Remove-Item $archive -Force; Remove-Item $tmp -Recurse -Force
        Write-OK "StaxRip installed"
        $results['StaxRip'] = "installed"
    } catch { Write-Fail "StaxRip: $($_.Exception.Message)"; $results['StaxRip'] = "FAILED" }
}

# --- comics-dl ---
$dest = Join-Path $root "core\comics-dl.exe"
if ((Test-Path $dest) -and -not $Force) {
    Write-Skip "comics-dl already installed"
    $results['comics-dl'] = "skipped (exists)"
} else {
    try {
        $rel   = Get-GitHubLatest "Girbons/comics-downloader"
        $asset = $rel.assets | Where-Object { $_.name -match 'windows.*amd64.*\.exe$' -or $_.name -match 'windows-amd64' } | Select-Object -First 1
        if (-not $asset) { $asset = $rel.assets | Where-Object { $_.name -match '\.exe$' } | Select-Object -First 1 }
        if (-not $asset) { throw "No Windows exe found in release" }
        Download-File $asset.browser_download_url $dest "comics-dl $($rel.tag_name)"
        Write-OK "comics-dl installed"
        $results['comics-dl'] = "installed"
    } catch { Write-Fail "comics-dl: $($_.Exception.Message)"; $results['comics-dl'] = "FAILED" }
}

# ----------------------------------------------------------------
# 4. Summary
# ----------------------------------------------------------------
Write-Host ""
Write-Host "[4/4] Setup complete" -ForegroundColor Yellow
Write-Host ""
Write-Host "================================================================" -ForegroundColor White
Write-Host "  Results" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor White
foreach ($key in $results.Keys) {
    $val   = $results[$key]
    $color = if ($val -like 'FAILED*') { 'Red' } elseif ($val -like 'skipped*') { 'Gray' } else { 'Green' }
    Write-Host ("  {0,-20} {1}" -f $key, $val) -ForegroundColor $color
}
Write-Host ""
Write-Host "  NOTE: HandBrake presets (presets.json, 480p.json, 2160p.json)" -ForegroundColor Yellow
Write-Host "        must be present in core\ for the encode scripts to work." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Logs will be written to: $root\logs\" -ForegroundColor White
Write-Host "  Drop source files into:  $root\Input\" -ForegroundColor White
Write-Host "  Encoded output goes to:  $root\Output\" -ForegroundColor White
Write-Host ""
