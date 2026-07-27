# ProAV Shoko

**ProAV Shoko** - A platform-independent analysis tool for AV environments.
Analyze, verify and troubleshoot USB connections in meeting rooms and BYOD environments.


[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/klangche/proav-shoko/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux%20%7C%20PowerShell-lightgrey)](https://github.com/klangche/proav-shoko)

---

## Table of Contents

- [Overview](#overview)
- [Why ProAV Shoko?](#why-proav-shoko)
- [Features](#features)
- [Downloads](#downloads)
- [5 Ways to Run](#5-ways-to-run)
- [Usage](#usage)
- [Tech Stack](#tech-stack)
- [Building From Source](#building-from-source)
- [License](#license)

---

## Overview

ProAV Shoko is a **platform-independent** analysis tool that:

- **Scans** all connected USB devices
- **Builds** a hierarchical tree of the USB chain
- **Calculates** hops (number of levels) and tiers (depth)
- **Assesses** stability based on the chain length
- **Shows** connected displays with resolution
- **Generates** professional HTML and PDF reports

**Perfect for:** AV technicians, IT support, sales, and diagnostics teams who need to quickly identify USB issues in conference rooms.

---

## Why ProAV Shoko?

In modern meeting rooms we often see:

- USB-C docks (Unisynk, HP, Lenovo, CalDigit, Logitech, TiGHT, Hyper, Targus...)
- Multiple hubs and/or active cables chained together adding tiers and hops.
- Videobars, speakerphones, touch panels, wireless presentation dongles, external drives etc.
- Users' own iPads/iPhones/Android/other devices

**The problem:** Long chains often cause issues **only on Apple Silicon Macs** (M1/M2/M3/M4), while Windows and Intel Macs typically work flawlessly.

**ProAV Shoko** helps technicians prove:
*"The chain has 5 hops → Windows & Intel OK, but Apple Silicon is not stable"*

---

## Features

| Feature | Description | Status |
|---------|-------------|--------|
| USB Tree | Hierarchical view of all connected devices | ✅ |
| Hops & Tiers | Calculates the number of levels and maximum depth | ✅ |
| Stability Assessment | Color-coded based on chain length | ✅ |
| Apple Silicon Warning | Special warning at 5+ hops | ✅ |
| Display Information | Shows connected displays with resolution | ✅ |
| HTML Report | Dark background, identical to the terminal | ✅ |
| PDF Report | Long, continuous page for printing | ✅ |
| GUI | Live overview with tree and log | ✅ |

---

## Downloads

> All pre-built binaries are available on the **[Releases page](https://github.com/klangche/klangche-proav-shoko/releases)**.
>
> No installation required — download, unzip, and run.

---

## 5 Ways to Run

### 1. Windows Portable (`.exe`)
```powershell
# CLI mode
proav-shoko-windows.exe --cli

# GUI mode (no arguments)
proav-shoko-windows.exe
```
**Portable** — no install needed, runs on any Windows 10/11 x64.

### 2. macOS App (`.app`)
```bash
# Intel Mac:
open ProAV\ Shoko.app

# Apple Silicon Mac (M1/M2/M3/M4):
open ProAV\ Shoko.app

# CLI mode (both architectures):
./proav-shoko-macos --cli
```
**Portable** — drag to Applications or run from anywhere.

### 3. Linux Executable
```bash
chmod +x proav-shoko-linux
./proav-shoko-linux --cli

# GUI mode (requires tkinter):
./proav-shoko-linux
```

### 4. PowerShell Script
```powershell
# Run directly from GitHub (always latest version):
irm https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.ps1 | iex
```
**No Python required** — runs on any Windows machine with PowerShell 5.1+.

### 5. Python Script (Source)
```bash
git clone https://github.com/klangche/klangche-proav-shoko
cd klangche-proav-shoko
pip install -e .
python run.py --cli    # CLI mode
python run.py          # GUI mode
```
**Full control** — works on Windows, macOS, and Linux.

---

## Usage

### CLI Mode
```bash
proav-shoko --cli
```

**Output includes:**
1. Platform info (OS, architecture, admin status)
2. Full USB tree with hops/tiers
3. Overall stability rating
4. Per-port stability (EXTERNAL / INTERNAL)
5. Connected displays with resolution
6. **Interactive monitoring** - press Enter to stop, then:
   - Choose report format: `[Enter]HTML / [P]DF / [N]o report`
   - Report auto-opens in browser/PDF viewer

### GUI Mode
```bash
proav-shoko
# Or
python run.py
```

**Features:**
- Live USB tree with stability
- Real-time connect/disconnect log
- Report generation with format selection
- Export CSV limits

### Report Output
Reports include:
- Full USB tree with Mermaid diagrams
- Per-port stability assessment
- Monitoring log (if monitoring was run)
- ⚠ Unstable devices detected during monitoring
- Platform-specific stability limits
- Connected displays
- Platform notes

---

## Tech Stack

- **Python 3.10+** - Core language
- **usbmonitor** - USB device monitoring
- **screeninfo** - Display information
- **weasyprint** - PDF generation
- **tkinter** - GUI (built-in)
- **PyInstaller** - Cross-platform builds
- **Mermaid.js** - Diagrams in HTML reports
- **PowerShell** - Windows launcher

---

## Building From Source

### Prerequisites
```bash
# macOS
brew install python cairo pango gdk-pixbuf libffi

# Linux (Ubuntu/Debian)
sudo apt-get install python3-dev libcairo2-dev libpango1.0-dev libgdk-pixbuf-2.0-dev libffi-dev

# Windows
# Install Python from python.org
# Install Visual C++ Build Tools
```

### Quick Build (Current Platform)
```bash
pip install -U pip pyinstaller
pip install -e .
python -m PyInstaller proav-shoko.spec
```

### Cross-Platform Build (via GitHub Actions)
The repository includes `.github/workflows/build.yml` that builds for all platforms on push to main.

### Output
```
dist/
├── proav-shoki.exe              # Windows CLI
├── ProAV Shoko.exe              # Windows GUI
├── proav-shoki-macos-intel      # macOS Intel CLI
├── ProAV Shoko.app              # macOS Intel GUI
├── proav-shoki-macos-arm64      # macOS Apple Silicon CLI
├── ProAV Shoko.app              # macOS Apple Silicon GUI
└── proav-shoki-linux-x86_64     # Linux CLI
```

---

## Project Structure

```
klangche-proav-shoko/
├── .github/workflows/build.yml   # CI/CD builds
├── src/
│   ├── main.py                   # Entry point (GUI/CLI)
│   ├── main_cli.py               # CLI logic
│   ├── gui.py                    # Tkinter GUI
│   ├── usb_analyzer.py           # USB tree, hops, stability
│   ├── display_analyzer.py       # Display detection
│   ├── report_generator.py       # HTML/PDF reports
│   ├── platform_utils.py         # Platform detection
│   └── assets/
│       ├── report.css            # Report styling
│       └── usb_data.csv          # Platform stability limits
├── run.py                        # Simple entry point
├── proav-shoko.spec              # PyInstaller spec
├── proav-shoko.ps1               # PowerShell launcher
├── proav-shoko_powershell.ps1    # Full PS implementation
├── build.py                      # Cross-platform build script
├── pyproject.toml
├── README.md
└── LICENSE
```

---

## License

MIT License - see [LICENSE](LICENSE) for details.

---

*Built for the ProAV community by Klangche*