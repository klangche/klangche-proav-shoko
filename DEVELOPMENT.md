# Development

## How to Run

### 1. Portable Executable (.exe) — Windows

Download the latest release from the **[Releases page](https://github.com/klangche/klangche-proav-shoko/releases)**.  
Unzip and run `ProAV Shoko <version>.exe`.

- Double-click → GUI mode
- Run from terminal with `--cli` → CLI mode

No installation or Python required.

### 2. Python (run from source) — any platform

Easiest with [`uv`](https://docs.astral.sh/uv/) (downloads its own Python, so system
Python 3.9 on macOS is fine):

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
source $HOME/.local/bin/env
uv run python run.py         # GUI mode
uv run python run.py --cli   # CLI mode
```

Or with pip (requires Python ≥ 3.10):

```bash
pip install -e .
proav-shoko --cli
```

Or run directly from a clone:

```bash
python run.py         # GUI mode
python run.py --cli   # CLI mode
```

### 3. PowerShell Script (.ps1) — Windows

```powershell
iex (irm https://raw.githubusercontent.com/klangche/klangche-proav-shoko/main/proav-shoko.ps1)
```

No Python required — runs on any Windows machine with PowerShell 5.1+.

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
```

**Features:**
- Live USB tree with stability
- Real-time connect/disconnect log
- Report generation with format selection

### Report Output
Reports include:
- Full USB tree with Mermaid diagrams
- Per-port stability assessment
- Monitoring log (if monitoring was run)
- Unstable devices detected during monitoring
- Platform-specific stability limits
- Connected displays
- Platform notes

## Building From Source (Windows)

```bash
pip install -e .[dev]
pip install pyinstaller
python build.py --build
```

The compiled `.exe` files will be in the `dist/` folder.
