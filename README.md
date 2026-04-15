# Media Scripts 2026

A portable, self-contained Windows media toolkit. Run `Setup.bat` once on a fresh machine and the entire environment — folders, scripts, and tools — is built automatically.

Also available as a **Python AIO edition** (see [Python Fork](#python-fork--media-scripts-2026-aio) below) that compiles into a standalone `media_scripts_setup.exe` with no dependencies.

---

## Quick Start

1. Place the `Media_Scripts_2026` folder anywhere you like.
2. Double-click **`Setup.bat`**.
3. Press **Enter** at the prompt for a full setup, or type flags (see below) and press Enter.
4. Wait for all tools to download. When complete, every bat file and tool will be ready to use.

> **Moved the folder?** Run `Setup_Path_Variables_If_Error.bat` to re-patch the portable Python paths.

---

## Setup.bat

A simple launcher that calls `Setup.ps1` via PowerShell with execution policy bypass. You can also pass flags directly at the prompt it displays before running.

```
Setup.bat
```

When opened, it shows the available flags and prompts for input. Press **Enter** to run a full setup with no flags.

---

## Setup.ps1

The main setup script. Safe to re-run — existing files and already-installed tools are skipped unless `-Force` is used.

### Parameters

| Flag | Description |
|---|---|
| _(none)_ | Full setup — directories, bat files, and all tools |
| `-Force` | Re-download all tools and overwrite existing bat files |
| `-DirectoriesOnly` | Create folders only, skip bat files and downloads |
| `-SkipBatFiles` | Skip writing bat files (tools still download) |
| `-SkipHandBrake` | Skip HandBrake download (~65 MB) |
| `-SkipStaxRip` | Skip StaxRip download |
| `-SkipTsMuxer` | Skip tsMuxer / tsMuxerGUI download |

### Running directly from PowerShell

```powershell
# Full setup
.\Setup.ps1

# Skip large downloads for a quick re-run
.\Setup.ps1 -SkipHandBrake -SkipStaxRip -SkipTsMuxer

# Re-download everything and overwrite all bat files
.\Setup.ps1 -Force

# Directories only
.\Setup.ps1 -DirectoriesOnly
```

---

## What Gets Created

### Directory Structure

```
Media_Scripts_2026\
├── Input\
│   └── done\          # Processed input files are moved here
├── Output\
│   ├── Music\
│   └── Comics\
├── logs\              # Timestamped encode logs
└── core\
    ├── ffmpeg\
    ├── mpv\
    ├── staxrip\
    └── tsmuxer\
```

### Bat Files Written

| File | Description |
|---|---|
| `Standard x264 Base Encode (HandBrake).bat` | Standard x264 encode via HandBrake, MP4 output |
| `Standard x264 Base Encode (HandBrake MKV).bat` | Standard x264 encode via HandBrake, MKV output |
| `Encode 480p Downscale (HandBrake).bat` | Downscale to 480p via HandBrake, MP4 output |
| `Encode 2160p Upscale (HandBrake).bat` | Upscale to 2160p via HandBrake, MP4 output |
| `Encode 480p Downscale (FFmpeg).bat` | Downscale to 480p via FFmpeg x264, MP4 output |
| `Encode 2160p Upscale (FFmpeg).bat` | Upscale to 2160p via FFmpeg x265, MP4 output |
| `Encode 480p Downscale (FFmpeg MKV).bat` | Downscale to 480p via FFmpeg, MKV with subtitle passthrough |
| `Encode 2160p Upscale (FFmpeg MKV).bat` | Upscale to 2160p via FFmpeg, MKV with subtitle passthrough |
| `SBS to Anaglyph 3D.bat` | Convert Half Side-by-Side 3D to Red/Cyan Dubois anaglyph |
| `Stream-DL.bat` | Download video via yt-dlp (saved as `youtube-dl.exe`) |
| `Spotify-Downloader.bat` | Download Spotify tracks via spotdl |
| `Download-Comics.bat` | Download comics via comics-dl |
| `Media_Rename_Only.bat` | Rename media in place using mnamer |
| `Open MPV.bat` | Launch MPV player with file picker or drag-and-drop |
| `Open StaxRip.bat` | Launch StaxRip GUI encoder |
| `Open tsMuxer.bat` | Launch tsMuxerGUI (falls back to CLI if GUI not present) |
| `FFmpeg Prompt.bat` | Open a command prompt with FFmpeg on the PATH |
| `Update-Tools.bat` | Check versions and update any installed tool |
| `Setup_Path_Variables_If_Error.bat` | Re-patch portable Python paths after moving the folder |

> All encode bat files support **drag-and-drop** (drop files directly onto the bat icon) as well as the **`Input\` directory** workflow.

### Tools Downloaded

| Tool | Source | Notes |
|---|---|---|
| **yt-dlp** | `yt-dlp/yt-dlp` (GitHub Releases) | Saved as `youtube-dl.exe` for legacy compatibility |
| **spotdl** | `spotDL/spotify-downloader` (GitHub Releases) | Spotify track downloader |
| **FFmpeg** | `GyanD/codexffmpeg` (GitHub Releases) | Full build with all codecs |
| **MPV** | `zhongfly/mpv-winbuild` (GitHub Releases) | Portable media player |
| **HandBrakeCLI** | `HandBrake/HandBrake` (GitHub Releases) | ~65 MB — skip with `-SkipHandBrake` |
| **StaxRip** | `staxrip/staxrip` (GitHub Releases) | GUI encoder — skip with `-SkipStaxRip` |
| **tsMuxer** | `justdan96/tsMuxer` (GitHub Releases) | TS/Blu-ray muxer CLI — skip with `-SkipTsMuxer` |
| **tsMuxerGUI** | `justdan96/tsMuxer` (GitHub Releases) | GUI front-end for tsMuxer — installed alongside CLI |
| **comics-dl** | `Girbons/comics-downloader` (GitHub Releases) | Comic book downloader |
| **7zr.exe** | `7-zip.org` | Standalone 7-Zip (~400 KB), used to extract `.7z` archives |
| **pycore** | `pasiegel/Media-Scripts-2026` (this repo) | Portable Python 3.12.3 + mnamer |

All tools are downloaded to their latest release version at time of setup. Run `Update-Tools.bat` at any time to check for and install updates.

### Optional Tools (Update-Tools only)

These are not downloaded during setup but can be installed or updated at any time via `Update-Tools.bat`.

| Tool | Option | Notes |
|---|---|---|
| **Rickinator** | `[9]` in Update-Tools | Download manager — reads URLs from `links.txt`, saves to a download folder. Extracted to the toolkit root. |

---

## Portable Python & mnamer

The `pycore` package installs a portable WinPython environment into `core\python-3.12.3\` along with the `mnamer` media renamer tool. After extraction, Setup automatically patches the Python paths for the current folder location so mnamer works without any system Python installation.

If you move the toolkit to a new location, run:

```
Setup_Path_Variables_If_Error.bat
```

This re-runs the WinPython path fix (`make_winpython_fix.bat`) and restores full mnamer functionality.

---

## Python Fork — Media Scripts 2026 AIO

A standalone Python implementation lives in `media-scripts-python/`. It mirrors all functionality of `Setup.ps1`, adds an interactive startup prompt, and compiles into a single portable `media_scripts_setup.exe` using PyInstaller — no Python installation required on the target machine.

### Files

| File | Description |
|---|---|
| `setup.py` | Full Python port — stdlib only, no pip dependencies |
| `media_scripts_setup.exe` | Compiled standalone executable (ready to run) |
| `Setup.bat` | Launcher — runs `media_scripts_setup.exe` if present, otherwise `python setup.py` |
| `Setup.ps1` | Copy of the PowerShell version, kept in sync |
| `build-exe.bat` | Recompiles `setup.py` into `media_scripts_setup.exe` using PyInstaller |

### Usage

**Run the compiled exe** (no Python required):
```
media_scripts_setup.exe
```

**Run as Python script** (requires Python 3.10+):
```
Setup.bat
```
or directly:
```
python setup.py
python setup.py --skip-handbrake --skip-staxrip
python setup.py --force
python setup.py --directories-only
```

Both the exe and the script show an interactive startup prompt when launched with no arguments — displaying available flags and waiting for input before proceeding.

**Recompile the exe:**
```
build-exe.bat
```
Installs PyInstaller if needed, compiles `setup.py`, and places `media_scripts_setup.exe` in the project root. `Setup.bat` auto-detects and uses it on next run.

### Parameters (Python / exe)

| Flag | Description |
|---|---|
| _(none)_ | Full setup — directories, bat files, and all tools |
| `--force` | Re-download all tools and overwrite existing bat files |
| `--directories-only` | Create folders only, skip bat files and downloads |
| `--skip-bat-files` | Skip writing bat files (tools still download) |
| `--skip-handbrake` | Skip HandBrake download (~65 MB) |
| `--skip-staxrip` | Skip StaxRip download |
| `--skip-tsmuxer` | Skip tsMuxer / tsMuxerGUI download |

### PyInstaller compile command (manual)

```
pip install pyinstaller
pyinstaller --onefile --name media_scripts_setup setup.py
```

The compiled `media_scripts_setup.exe` is fully self-contained and runs on any Windows 10/11 machine without Python installed.

---

## Requirements

- Windows 10 / 11
- PowerShell 5.1 or later (included with Windows)
- Internet connection for initial tool downloads
- No admin rights required

> **Python AIO edition only:** Python 3.10+ required to run `setup.py` directly. Not required if using the compiled `media_scripts_setup.exe`.
