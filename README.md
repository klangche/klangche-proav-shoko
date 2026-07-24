# 🔍 ProAV Shōko

**ProAV Shōko** – Your platform-independent USB detective for AV environments.
Analyze, verify and troubleshoot USB connections in meeting rooms and BYOD environments.

[![Build Status](https://github.com/klangche/proav-shoko/actions/workflows/build.yml/badge.svg)](https://github.com/klangche/proav-shoko/actions)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/klangche/proav-shoko/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](https://github.com/klangche/proav-shoko)

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Why ProAV Shōko?](#-why-proav-shōko)
- [Features](#-features)
- [Installation](#-installation)
- [Usage](#-usage)
- [Tech Stack](#-tech-stack)
- [Building From Source](#-building-from-source)
- [License](#-license)

---

## 🎯 Overview

ProAV Shōko is a **platform-independent** analysis tool that:

- 📡 **Scans** all connected USB devices
- 🌳 **Builds** a hierarchical tree of the USB chain
- 📊 **Calculates** hops (number of levels) and tiers (depth)
- 🟢🟡🔴 **Assesses** stability based on the chain length
- 🖥️ **Shows** connected displays with resolution
- 📄 **Generates** professional HTML and PDF reports

> **Perfect for:** AV technicians, IT support, sales, and diagnostics teams who need to quickly identify USB issues in conference rooms.

---

## 🤔 Why ProAV Shōko?

In modern meeting rooms we often see:

- USB-C docks (Unisynk, HP, Lenovo, CalDigit, Logitech, TiGHT, Hyper, Targus...)
- Multiple hubs chained together
- Webcams, speakerphones, touch panels, wireless presentation dongles, external drives
- Users' own iPads/iPhones/Android devices

**The problem:** Long chains often cause issues **only on Apple Silicon Macs** (M1/M2/M3/M4), while Windows and Intel Macs typically work flawlessly.

**ProAV Shōko** helps technicians prove:
→ *"The chain has 5 hops → Windows & Intel OK, but Apple Silicon is not stable"*

---

## ✨ Features

| Feature | Description | Status |
|---------|-------------|--------|
| 🌳 **USB Tree** | Hierarchical view of all connected devices | ✅ |
| 📏 **Hops & Tiers** | Calculates the number of levels and maximum depth | ✅ |
| 🟢🟡🔴 **Stability Assessment** | Color-coded based on chain length | ✅ |
| 🍎 **Apple Silicon Warning** | Special warning at 5+ hops | ✅ |
| 🖥️ **Display Information** | Shows connected displays with resolution | ✅ |
| 📄 **HTML Report** | Dark background, identical to the terminal | ✅ |
| 📄 **PDF Report** | Long, continuous page for printing | ✅ |
| 🖥️ **GUI** | Live overview with tree and log | ✅ |

---

## 🚀 Installation

### 📦 For Users (executable file)

Download the latest version for your operating system from [Releases](https://github.com/klangche/proav-shoko/releases):

| Platform | File | Size |
|----------|------|------|
| 🪟 **Windows** | `proav-shoko-windows.exe` | ~15 MB |
| 🍎 **macOS (Intel)** | `proav-shoko-macos-intel` | ~18 MB |
| 🍎 **macOS (Apple Silicon)** | `proav-shoko-macos-arm64` | ~18 MB |
| 🐧 **Linux** | `proav-shoko-linux` | ~15 MB |

```bash
# Example: run directly on macOS
chmod +x proav-shoko-macos-arm64
./proav-shoko-macos-arm64
```
