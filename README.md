# ProAV Shoko

**ProAV Shoko** - A platform-independent analysis tool for AV environments.
Analyze, verify and troubleshoot USB connections in meeting rooms and BYOD environments.

[![Releases](https://img.shields.io/github/v/release/klangche/klangche-proav-shoko?label=download)](https://github.com/klangche/klangche-proav-shoko/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux%20%7C%20Python%20%7C%20PowerShell-lightgrey)](https://github.com/klangche/klangche-proav-shoko)
[![SemVer](https://img.shields.io/badge/versioning-SemVer%202.0.0-blue)](VERSIONING.md)

---

## Quick Start

Download the latest release for your platform:

| Platform | Download |
|----------|----------|
| **Windows, MacOS, Linux** | [`ProAV Shoko`](https://github.com/klangche/klangche-proav-shoko/releases) — Download and run!

**Run from source — no Python, no git, no admin required.**

Paste one line into Terminal or PowerShell. [`uv`](https://docs.astral.sh/uv/) downloads its own Python and installs to a temporary user cache.

### CLI version — Terminal / PowerShell

*Mac / Linux:*
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh && ~/.local/bin/uvx --python 3.12 --from "https://github.com/klangche/klangche-proav-shoko/archive/refs/heads/main.zip" proav-shoko
```

*Windows PowerShell:*
```powershell
irm https://astral.sh/uv/install.ps1 | iex; & "$env:USERPROFILE\.local\bin\uvx.exe" --python 3.12 --from "https://github.com/klangche/klangche-proav-shoko/archive/refs/heads/main.zip" proav-shoko
```

### GUI version — beautiful GUI

*Mac / Linux:*
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh && ~/.local/bin/uvx --python 3.12 --from "https://github.com/klangche/klangche-proav-shoko/archive/refs/heads/main.zip" proav-shoko --gui
```

*Windows PowerShell:*
```powershell
irm https://astral.sh/uv/install.ps1 | iex; & "$env:USERPROFILE\.local\bin\uvx.exe" --python 3.12 --from "https://github.com/klangche/klangche-proav-shoko/archive/refs/heads/main.zip" proav-shoko --gui
```

> **pip alternative** — requires Python ≥ 3.10, no git needed:
> `python3 -m pip install --target /tmp/proav-shoko "https://github.com/klangche/klangche-proav-shoko/archive/refs/heads/main.zip" && PYTHONPATH=/tmp/proav-shoko python3 -m src`



## Table of Contents

- [Overview](#overview)
- [Why ProAV Shoko?](#why-proav-shoko)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Versioning](VERSIONING.md)
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

## Tech Stack

- **Python 3.10+** - Core language
- **usbmonitor** - USB device monitoring
- **screeninfo** - Display information
- **weasyprint** - PDF generation (bundled in packaged apps)
- **customtkinter** - Modern GUI theme (includes tkinter)
- **PyInstaller** - Cross-platform builds (Windows exe, macOS Universal, Linux)
- **Mermaid.js** - Diagrams in HTML reports
- **PowerShell** - Windows launcher

---

## Project Structure

```
klangche-proav-shoko/
├── src/
│   ├── main.py                   # Entry point (GUI/CLI)
│   ├── main_cli.py               # CLI logic
│   ├── gui.py                    # Tkinter GUI
│   ├── usb_analyzer.py           # USB tree, hops, stability
│   ├── usb_topology.py           # Platform-specific parent detection
│   ├── display_analyzer.py       # Display detection
│   ├── report_generator.py       # HTML/PDF reports
│   ├── platform_utils.py         # Platform detection
│   ├── resources/
│   │   ├── shoko-icon.png        # App icon (source)
│   │   └── shoko-icon.ico        # Auto-generated from .png
│   └── assets/
│       ├── report.css            # Report styling
│       └── usb_data.csv          # Platform stability limits
├── scripts/
│   └── generate_icon.py          # Auto-generates .ico from .png
├── proav-shoko.ps1               # PowerShell launcher
├── proav-shoko_powershell.ps1    # Full PS implementation
├── run.py                        # Python entry point
├── build.py                      # Build script
├── proav-shoko.spec              # PyInstaller spec
├── pyproject.toml
├── proav-shoko.json              # Configuration
├── README.md
├── LICENSE
└── .github/workflows/build.yml   # CI
```

---

## License

MIT License - see [LICENSE](LICENSE) for details.

---

*Built for the ProAV community by Klangche*
