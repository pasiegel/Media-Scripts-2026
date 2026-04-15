#!/usr/bin/env python3
"""
Media Scripts 2026 - Setup (Python edition)
Portable Windows media toolkit installer.
Compile to exe with: pyinstaller --onefile setup.py
"""

import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import urllib.request
import urllib.error
import zipfile
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Path resolution — works both as .py and as a PyInstaller frozen .exe
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).parent

# ---------------------------------------------------------------------------
# Console colour helpers (ANSI; enable virtual terminal on Windows 10+)
# ---------------------------------------------------------------------------
def _enable_ansi():
    if sys.platform == 'win32':
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
        except Exception:
            pass

_enable_ansi()

CYAN   = '\033[96m'
GREEN  = '\033[92m'
YELLOW = '\033[93m'
RED    = '\033[91m'
GRAY   = '\033[90m'
RESET  = '\033[0m'

def step(msg):  print(f"  {CYAN}{msg}{RESET}")
def ok(msg):    print(f"  {GREEN}[OK]  {msg}{RESET}")
def skip(msg):  print(f"  {GRAY}[--]  {msg}{RESET}")
def fail(msg):  print(f"  {RED}[ERR] {msg}{RESET}")
def warn(msg):  print(f"  {YELLOW}[!]   {msg}{RESET}")

# ---------------------------------------------------------------------------
# Download helpers
# ---------------------------------------------------------------------------
def download(url: str, dest: Path, label: str = ''):
    label = label or dest.name
    step(f"Downloading {label}...")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'media-scripts-setup/1.0'})
        with urllib.request.urlopen(req) as resp, open(dest, 'wb') as f:
            shutil.copyfileobj(resp, f)
    except urllib.error.URLError as e:
        raise RuntimeError(f"Download failed: {e}") from e


def github_latest(repo: str) -> dict:
    """Return the latest GitHub release JSON for owner/repo."""
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(url, headers={'User-Agent': 'media-scripts-setup/1.0'})
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def find_asset(release: dict, pattern: str) -> dict | None:
    """Case-insensitive substring match against asset names."""
    import re
    pat = re.compile(pattern, re.IGNORECASE)
    for asset in release.get('assets', []):
        if pat.search(asset['name']):
            return asset
    return None


def extract_zip(zip_path: Path, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(dest)


def get_7zr() -> Path:
    """Return path to 7zr.exe, downloading if missing."""
    p = ROOT / 'core' / '7zr.exe'
    if not p.exists():
        step("Downloading 7zr.exe (~400 KB)...")
        download('https://www.7-zip.org/a/7zr.exe', p, '7zr.exe')
    return p


def extract_7z(archive: Path, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    zr = get_7zr()
    subprocess.run([str(zr), 'x', str(archive), f'-o{dest}', '-y'],
                   check=True, stdout=subprocess.DEVNULL)


def flatten_single_subdir(src: Path, dest: Path):
    """If src contains exactly one subdirectory, copy its contents to dest."""
    children = list(src.iterdir())
    sub = None
    if len(children) == 1 and children[0].is_dir():
        sub = children[0]
    src_dir = sub if sub else src
    for item in src_dir.iterdir():
        target = dest / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True)
        else:
            shutil.copy2(item, target)

# ---------------------------------------------------------------------------
# Bat file content — single source of truth
# ---------------------------------------------------------------------------
BAT_FILES: dict[str, str] = {}

BAT_FILES['Download-Comics.bat'] = r"""@echo off
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
"""

BAT_FILES['Media_Rename_Only.bat'] = (
    r'.\core\python-3.12.3\Scripts\mnamer.exe'
    r' --movie-directory="{name} ({year})" --episode-directory="{series}" .'
    '\n'
)

BAT_FILES['Setup_Path_Variables_If_Error.bat'] = r'.\core\scripts\make_winpython_fix.bat' + '\n'

BAT_FILES['Spotify-Downloader.bat'] = r"""@echo off
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
"""

BAT_FILES['Stream-DL.bat'] = r"""@echo off
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
"""

BAT_FILES['Standard x264 Base Encode (HandBrake).bat'] = r"""@echo off
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
"""

BAT_FILES['Standard x264 Base Encode (HandBrake MKV).bat'] = r"""@echo off
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
"""

BAT_FILES['Encode 480p Downscale (HandBrake).bat'] = r"""@echo off
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
"""

BAT_FILES['Encode 2160p Upscale (HandBrake).bat'] = r"""@echo off
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
"""

BAT_FILES['Encode 480p Downscale (FFmpeg).bat'] = r"""@echo off
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
"""

BAT_FILES['Encode 2160p Upscale (FFmpeg).bat'] = r"""@echo off
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
"""

BAT_FILES['Encode 480p Downscale (FFmpeg MKV).bat'] = r"""@echo off
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
echo  Subtitles: all tracks copied (ASS/SRT/PGS/VOBSUB supported)
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
"""

BAT_FILES['Encode 2160p Upscale (FFmpeg MKV).bat'] = r"""@echo off
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
"""

BAT_FILES['SBS to Anaglyph 3D.bat'] = r"""@echo off
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
"""

BAT_FILES['FFmpeg Prompt.bat'] = r"""@echo off
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
"""

BAT_FILES['Open MPV.bat'] = r"""@echo off
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
"""

BAT_FILES['Open StaxRip.bat'] = r"""@echo off
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
"""

BAT_FILES['Open tsMuxer.bat'] = r"""@echo off
@title tsMuxer
setlocal
cd /d "%~dp0"

set TSMUXER=.\core\tsmuxer\tsMuxer.exe
set TSMUXERGUI=.\core\tsmuxer\tsMuxerGUI.exe

if exist "%TSMUXERGUI%" (
    echo Launching tsMuxerGUI...
    start "" "%TSMUXERGUI%"
    goto :end
)

if exist "%TSMUXER%" (
    echo tsMuxerGUI not found - launching CLI tsMuxer instead.
    echo Usage: tsMuxer.exe <metafile> <output>
    start "" cmd /k "cd /d "%~dp0core\tsmuxer" && echo tsMuxer CLI ready. Type tsMuxer.exe for usage."
    goto :end
)

echo tsMuxer not found at .\core\tsmuxer\
echo Run Update-Tools.bat and select tsMuxer to download automatically.

:end
pause
endlocal
"""

# core helper bats (written to core\ subdirectory)
CORE_BAT_FILES: dict[str, str] = {
    'rename.bat': (
        r'.\core\python-3.12.3\Scripts\mnamer.exe'
        r' --batch --no-overwrite'
        r' --movie-directory="{name} ({year})"'
        r' --episode-directory="{series}" output'
        '\n'
    ),
    'mnamer.bat': (
        r'.\core\python-3.12.3\Scripts\mnamer.exe'
        r' --movie-directory="{name} ({year})" --episode-directory="{series}" .'
        '\n'
    ),
}

# ---------------------------------------------------------------------------
# Directory structure
# ---------------------------------------------------------------------------
DIRS = [
    'Input',
    'Input/done',
    'Output',
    'Output/Music',
    'Output/Comics',
    'logs',
    'core/ffmpeg',
    'core/mpv',
    'core/staxrip',
    'core/tsmuxer',
]

# ---------------------------------------------------------------------------
# Tool download functions
# ---------------------------------------------------------------------------

def install_ytdlp(force: bool = False):
    dest = ROOT / 'core' / 'youtube-dl.exe'
    if dest.exists() and not force:
        skip('yt-dlp (youtube-dl.exe) already installed')
        return
    print(f"\n  {YELLOW}[yt-dlp]{RESET} Downloading latest release...")
    try:
        url = 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe'
        step(f"Source: {url}")
        download(url, dest, 'yt-dlp.exe')
        ok("Saved as youtube-dl.exe (legacy compatibility name)")
    except Exception as e:
        fail(str(e))


def install_spotdl(force: bool = False):
    dest = ROOT / 'core' / 'spotdl.exe'
    if dest.exists() and not force:
        skip('spotdl already installed')
        return
    print(f"\n  {YELLOW}[spotdl]{RESET} Downloading latest release...")
    try:
        rel = github_latest('spotDL/spotify-downloader')
        asset = (find_asset(rel, r'^spotdl.*windows.*\.exe$') or
                 find_asset(rel, r'\.exe$'))
        if not asset:
            fail("Could not find Windows exe in release assets")
            return
        step(f"Downloading: {asset['name']}")
        download(asset['browser_download_url'], dest, asset['name'])
        ok(f"spotdl installed ({rel['tag_name']})")
    except Exception as e:
        fail(str(e))


def install_ffmpeg(force: bool = False):
    dest_dir = ROOT / 'core' / 'ffmpeg'
    ffmpeg_exe = dest_dir / 'ffmpeg.exe'
    if ffmpeg_exe.exists() and not force:
        skip('FFmpeg already installed')
        return
    print(f"\n  {YELLOW}[FFmpeg]{RESET} Downloading latest essentials build...")
    try:
        rel = github_latest('GyanD/codexffmpeg')
        step(f"Latest release: {rel['tag_name']}")
        asset = find_asset(rel, r'essentials_build.*\.zip$')
        if not asset:
            fail("No essentials zip found in release")
            return
        mb = round(asset['size'] / 1024 / 1024, 1)
        step(f"Downloading: {asset['name']} ({mb} MB)")
        zip_path = dest_dir / '_dl.zip'
        download(asset['browser_download_url'], zip_path, asset['name'])
        tmp = dest_dir / '_tmp'
        if tmp.exists():
            shutil.rmtree(tmp)
        extract_zip(zip_path, tmp)
        for exe in tmp.rglob('*.exe'):
            if exe.name in ('ffmpeg.exe', 'ffprobe.exe'):
                shutil.copy2(exe, dest_dir / exe.name)
                ok(f"Installed: {exe.name}")
        zip_path.unlink(missing_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception as e:
        fail(str(e))


def install_mpv(force: bool = False):
    dest_dir = ROOT / 'core' / 'mpv'
    mpv_exe = dest_dir / 'mpv.exe'
    if mpv_exe.exists() and not force:
        skip('MPV already installed')
        return
    print(f"\n  {YELLOW}[MPV]{RESET} Downloading latest Windows build...")
    try:
        rel = github_latest('zhongfly/mpv-winbuild')
        step(f"Latest release: {rel['tag_name']}")
        asset = find_asset(rel, r'mpv-x86_64.*\.7z$')
        if not asset:
            fail("No x64 .7z found in release")
            return
        mb = round(asset['size'] / 1024 / 1024, 1)
        step(f"Downloading: {asset['name']} ({mb} MB)")
        archive = dest_dir / '_dl.7z'
        download(asset['browser_download_url'], archive, asset['name'])
        step("Extracting...")
        tmp = dest_dir / '_tmp'
        if tmp.exists():
            shutil.rmtree(tmp)
        extract_7z(archive, tmp)
        for exe in tmp.rglob('mpv.exe'):
            shutil.copy2(exe, dest_dir / 'mpv.exe')
            ok("Installed: mpv.exe")
            break
        archive.unlink(missing_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception as e:
        fail(str(e))


def install_handbrake(force: bool = False, skip_flag: bool = False):
    if skip_flag:
        skip('HandBrake (skipped via --skip-handbrake)')
        return
    dest = ROOT / 'core' / 'HandBrakeCLI.exe'
    if dest.exists() and not force:
        skip('HandBrakeCLI already installed')
        return
    print(f"\n  {YELLOW}[HandBrakeCLI]{RESET} Downloading latest release (~65 MB)...")
    try:
        rel = github_latest('HandBrake/HandBrake')
        step(f"Latest release: {rel['tag_name']}")
        asset = find_asset(rel, r'HandBrakeCLI.*x86_64.*\.zip$')
        if not asset:
            fail("Could not find HandBrakeCLI zip. Check https://handbrake.fr/downloads2.php")
            return
        mb = round(asset['size'] / 1024 / 1024, 1)
        step(f"Downloading: {asset['name']} ({mb} MB)")
        zip_path = ROOT / 'core' / '_hb_dl.zip'
        download(asset['browser_download_url'], zip_path, asset['name'])
        tmp = ROOT / 'core' / '_hb_tmp'
        if tmp.exists():
            shutil.rmtree(tmp)
        extract_zip(zip_path, tmp)
        for exe in tmp.rglob('HandBrakeCLI.exe'):
            shutil.copy2(exe, dest)
            ok("HandBrakeCLI.exe installed")
            break
        zip_path.unlink(missing_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception as e:
        fail(str(e))


def install_staxrip(force: bool = False, skip_flag: bool = False):
    if skip_flag:
        skip('StaxRip (skipped via --skip-staxrip)')
        return
    dest_dir = ROOT / 'core' / 'staxrip'
    staxrip_exe = dest_dir / 'StaxRip.exe'
    if staxrip_exe.exists() and not force:
        skip('StaxRip already installed')
        return
    print(f"\n  {YELLOW}[StaxRip]{RESET} Downloading latest release...")
    try:
        rel = github_latest('staxrip/staxrip')
        step(f"Latest release: {rel['tag_name']}")
        asset = (find_asset(rel, r'StaxRip.*x64.*\.7z$') or
                 find_asset(rel, r'\.7z$'))
        if not asset:
            fail("No .7z asset found. Check https://github.com/staxrip/staxrip/releases")
            return
        mb = round(asset['size'] / 1024 / 1024, 1)
        step(f"Downloading: {asset['name']} ({mb} MB)")
        archive = dest_dir / '_dl.7z'
        download(asset['browser_download_url'], archive, asset['name'])
        step("Extracting...")
        tmp = dest_dir / '_tmp'
        if tmp.exists():
            shutil.rmtree(tmp)
        extract_7z(archive, tmp)
        flatten_single_subdir(tmp, dest_dir)
        archive.unlink(missing_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)
        ok("StaxRip installed to core/staxrip/")
    except Exception as e:
        fail(str(e))


def install_comicsdl(force: bool = False):
    dest = ROOT / 'core' / 'comics-dl.exe'
    if dest.exists() and not force:
        skip('comics-dl already installed')
        return
    print(f"\n  {YELLOW}[comics-dl]{RESET} Downloading latest release...")
    try:
        rel = github_latest('Girbons/comics-downloader')
        step(f"Latest release: {rel['tag_name']}")
        asset = (find_asset(rel, r'windows.*amd64.*\.exe$') or
                 find_asset(rel, r'windows-amd64') or
                 find_asset(rel, r'\.exe$'))
        if not asset:
            fail("Could not find Windows exe. Check https://github.com/Girbons/comics-downloader/releases")
            return
        step(f"Downloading: {asset['name']}")
        download(asset['browser_download_url'], dest, asset['name'])
        ok(f"Saved as comics-dl.exe ({rel['tag_name']})")
    except Exception as e:
        fail(str(e))


def install_7zr(force: bool = False):
    dest = ROOT / 'core' / '7zr.exe'
    if dest.exists() and not force:
        skip('7zr.exe already installed')
        return
    print(f"\n  {YELLOW}[7-Zip]{RESET} Downloading 7zr.exe...")
    try:
        download('https://www.7-zip.org/a/7zr.exe', dest, '7zr.exe')
        ok("7zr.exe installed to core/")
    except Exception as e:
        fail(str(e))


def install_tsmuxer(force: bool = False, skip_flag: bool = False):
    """Install both tsMuxer CLI and tsMuxerGUI from justdan96/tsMuxer."""
    if skip_flag:
        skip('tsMuxer (skipped via --skip-tsmuxer)')
        return
    dest_dir = ROOT / 'core' / 'tsmuxer'
    cli_exe = dest_dir / 'tsMuxer.exe'
    gui_exe = dest_dir / 'tsMuxerGUI.exe'
    if cli_exe.exists() and gui_exe.exists() and not force:
        skip('tsMuxer already installed')
        return
    print(f"\n  {YELLOW}[tsMuxer]{RESET} Downloading latest release...")
    try:
        rel = github_latest('justdan96/tsMuxer')
        tag = rel['tag_name']
        step(f"Latest release: {tag}")

        # tsMuxer CLI
        if not cli_exe.exists() or force:
            asset = find_asset(rel, r'tsMuxer.*win64.*\.zip$')
            if not asset:
                # Fallback: construct URL from tag
                asset_url = f"https://github.com/justdan96/tsMuxer/releases/download/{tag}/tsMuxer-{tag}-win64.zip"
                step(f"Downloading tsMuxer CLI from {asset_url}")
                zip_path = dest_dir / '_tsmuxer_dl.zip'
                download(asset_url, zip_path, f'tsMuxer-{tag}-win64.zip')
            else:
                step(f"Downloading: {asset['name']}")
                zip_path = dest_dir / '_tsmuxer_dl.zip'
                download(asset['browser_download_url'], zip_path, asset['name'])
            tmp = dest_dir / '_tsmuxer_tmp'
            if tmp.exists():
                shutil.rmtree(tmp)
            extract_zip(zip_path, tmp)
            for exe in tmp.rglob('tsMuxer.exe'):
                shutil.copy2(exe, dest_dir / 'tsMuxer.exe')
                ok("Installed: tsMuxer.exe")
                break
            zip_path.unlink(missing_ok=True)
            shutil.rmtree(tmp, ignore_errors=True)

        # tsMuxerGUI
        if not gui_exe.exists() or force:
            asset = find_asset(rel, r'tsMuxerGUI.*win64.*\.zip$')
            if not asset:
                asset_url = f"https://github.com/justdan96/tsMuxer/releases/download/{tag}/tsMuxerGUI-{tag}-win64.zip"
                step(f"Downloading tsMuxerGUI from {asset_url}")
                zip_path = dest_dir / '_tsmuxergui_dl.zip'
                download(asset_url, zip_path, f'tsMuxerGUI-{tag}-win64.zip')
            else:
                step(f"Downloading: {asset['name']}")
                zip_path = dest_dir / '_tsmuxergui_dl.zip'
                download(asset['browser_download_url'], zip_path, asset['name'])
            tmp = dest_dir / '_tsmuxergui_tmp'
            if tmp.exists():
                shutil.rmtree(tmp)
            extract_zip(zip_path, tmp)
            for exe in tmp.rglob('tsMuxerGUI.exe'):
                shutil.copy2(exe, dest_dir / 'tsMuxerGUI.exe')
                ok("Installed: tsMuxerGUI.exe")
                break
            zip_path.unlink(missing_ok=True)
            shutil.rmtree(tmp, ignore_errors=True)

    except Exception as e:
        fail(str(e))


def install_pycore(force: bool = False):
    """Download portable Python + mnamer from pasiegel/Media-Scripts-2026."""
    dest_dir = ROOT / 'core' / 'python-3.12.3'
    if dest_dir.exists() and not force:
        skip('pycore (portable Python + mnamer) already installed')
        return
    print(f"\n  {YELLOW}[pycore]{RESET} Downloading portable Python + mnamer...")
    try:
        rel = github_latest('pasiegel/Media-Scripts-2026')
        asset = find_asset(rel, r'pycore.*\.zip$') or find_asset(rel, r'\.zip$')
        if not asset:
            warn("pycore zip not found in release assets. Skipping.")
            return
        step(f"Downloading: {asset['name']}")
        zip_path = ROOT / 'core' / '_pycore_dl.zip'
        download(asset['browser_download_url'], zip_path, asset['name'])
        step("Extracting...")
        tmp = ROOT / 'core' / '_pycore_tmp'
        if tmp.exists():
            shutil.rmtree(tmp)
        extract_zip(zip_path, tmp)
        flatten_single_subdir(tmp, ROOT / 'core')
        zip_path.unlink(missing_ok=True)
        shutil.rmtree(tmp, ignore_errors=True)
        ok("pycore installed to core/")
    except Exception as e:
        fail(str(e))


# ---------------------------------------------------------------------------
# Bat file writer
# ---------------------------------------------------------------------------
def write_bat(rel_path: str, content: str, force: bool):
    path = ROOT / rel_path
    if path.exists() and not force:
        skip(rel_path)
    else:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding='ascii', errors='replace')
        ok(rel_path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def _interactive_prompt():
    """If no CLI args were given, show the options banner and prompt for flags.
    Any flags the user types are injected into sys.argv so argparse picks them up.
    Skipped automatically when args are already present (e.g. called from a script).
    """
    if len(sys.argv) > 1:
        return

    import shlex

    print()
    print('=' * 64)
    print('  Media Scripts 2026 - AIO')
    print('  Creates folders, writes bat files, and downloads all tools.')
    print('=' * 64)
    print()
    print('  Options (leave blank for full setup):')
    print('    --force             Re-download tools and overwrite bat files')
    print('    --directories-only  Create folders only, skip everything else')
    print('    --skip-bat-files    Skip writing bat files')
    print('    --skip-handbrake    Skip HandBrake (~65 MB)')
    print('    --skip-staxrip      Skip StaxRip')
    print('    --skip-tsmuxer      Skip tsMuxer / tsMuxerGUI')
    print()
    print('  NOTE: If you move this folder to a new location later, run')
    print('        Setup_Path_Variables_If_Error.bat to re-patch Python paths.')
    print()

    try:
        extra = input('  Extra flags (or press Enter for full setup): ').strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return

    if extra:
        sys.argv.extend(shlex.split(extra))


def main():
    _interactive_prompt()

    parser = argparse.ArgumentParser(
        description='Media Scripts 2026 - Setup (Python edition)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('--force',            action='store_true', help='Re-download tools and overwrite bat files')
    parser.add_argument('--directories-only', action='store_true', help='Create folders only, skip bat files and downloads')
    parser.add_argument('--skip-bat-files',   action='store_true', help='Skip writing bat files')
    parser.add_argument('--skip-handbrake',   action='store_true', help='Skip HandBrake download (~65 MB)')
    parser.add_argument('--skip-staxrip',     action='store_true', help='Skip StaxRip download')
    parser.add_argument('--skip-tsmuxer',     action='store_true', help='Skip tsMuxer download')
    args = parser.parse_args()

    print()
    print('=' * 64)
    print(f'  Media Scripts 2026 - Setup  (Python edition)')
    print(f'  Root: {ROOT}')
    print('=' * 64)

    # ------------------------------------------------------------------
    # 1. Directories
    # ------------------------------------------------------------------
    print(f'\n{YELLOW}[1/4] Creating directory structure...{RESET}')
    for d in DIRS:
        path = ROOT / d.replace('/', os.sep)
        if path.exists():
            skip(d)
        else:
            path.mkdir(parents=True, exist_ok=True)
            ok(f"Created: {d}")

    if args.directories_only:
        print(f'\n  {GREEN}Directory setup complete (--directories-only).{RESET}')
        return

    # ------------------------------------------------------------------
    # 2. Bat files
    # ------------------------------------------------------------------
    print(f'\n{YELLOW}[2/4] Writing bat files...{RESET}')
    if args.skip_bat_files:
        skip('Bat file creation skipped (--skip-bat-files)')
    else:
        for name, content in BAT_FILES.items():
            write_bat(name, content, args.force)
        for name, content in CORE_BAT_FILES.items():
            write_bat(f'core/{name}', content, args.force)

    # ------------------------------------------------------------------
    # 3. Download tools
    # ------------------------------------------------------------------
    print(f'\n{YELLOW}[3/4] Downloading tools...{RESET}')
    install_7zr(args.force)
    install_ytdlp(args.force)
    install_spotdl(args.force)
    install_ffmpeg(args.force)
    install_mpv(args.force)
    install_handbrake(args.force, skip_flag=args.skip_handbrake)
    install_staxrip(args.force, skip_flag=args.skip_staxrip)
    install_comicsdl(args.force)
    install_tsmuxer(args.force, skip_flag=args.skip_tsmuxer)
    install_pycore(args.force)

    # ------------------------------------------------------------------
    # 4. Done
    # ------------------------------------------------------------------
    print(f'\n{YELLOW}[4/4] Setup complete.{RESET}')
    print()
    print('=' * 64)
    print(f'  {GREEN}All done!{RESET}')
    print(f'  Root: {ROOT}')
    print('=' * 64)
    print()
    print('  If you move this folder, run:')
    print('    Setup_Path_Variables_If_Error.bat')
    print('  to re-patch portable Python paths.')
    print()


if __name__ == '__main__':
    main()
