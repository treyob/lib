# Chimera Toolbox

A PowerShell-based toolkit for Windows system maintenance, troubleshooting, and dental software fixes. A menu-driven terminal interface puts common IT admin tasks (network repairs, system scans, software installs, and dental imaging fixes) one keypress away.

---

## Quick Start

Run from any Windows PowerShell prompt:

```powershell
irm www.chimera-it.com/ps | iex
```

If you get a TLS/SSL error, force TLS 1.2 first:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
irm www.chimera-it.com/ps | iex
```

Alternatively, download `chimera-toolbox.exe` from the [Releases](https://github.com/treyob/lib/releases) page and double-click it.

The launcher downloads the toolbox to `%TEMP%\chimera-toolbox`, runs it in the current terminal, and deletes it on exit.

---

## Features

### System & Network

- Release and renew IP address
- Flush and register DNS
- Disable IPv6
- Repair Windows system files (SFC and DISM)
- Run CHKDSK on `C:\`
- Reset print spooler
- Reinstall Microsoft Print to PDF
- Run Group Policy update
- Get device serial number

### Administration & Remote Access

- Create a local admin user
- Enable PSRemoting
- Enable Remote Desktop
- Set firewall profiles to allow all

### Software & Utilities

- Install Microsoft Office
- Install Visual C++ Redistributables (2005 to v14, silent)
- .NET Repair Tool
- Network scanning (CLI scan and Advanced IP Scanner)
- Hardware and disk tools (HWInfo, USB Tree View, WizTree)
- Remote support (TeamViewer QuickSupport)
- Revo Uninstaller
- Hold music playback
- Games (Battleship, LibreMines)

### Dental Software

| Vendor | Tools |
| --- | --- |
| **Patterson (Eaglesoft)** | General fix, SmartDoc printer reinstall, Dexis/Gendex and Schick sensor integration, step-by-step IOSS configuration |
| **Vatech (EzDent)** | `libiomp5md.dll` error fix, clear EzDent cache |
| **Dentsply Sirona** | Schick drivers, Sidexis 4.3 migration |
| **Others** | Mouthwatch and Daryou camera drivers, TDO XDR sensor fix for Windows 11 |

#### Most Used

- Eaglesoft "unable to connect to server" fix
- Eaglesoft hexadecimal error fix
- Mouthwatch camera driver install

---

## Repository Layout

```text
lib/
├── build/                  # compile-public-chimera-toolbox.ps1 - builds chimera-toolbox.exe/.zip
├── src/
│   ├── app/                # Terminal UI (Python / Textual)
│   └── backend/
│       ├── data/           # Config and data files
│       ├── modules/        # PowerShell modules (general, dental, games)
│       └── scripts/        # Launcher fetched by the irm one-liner
└── requirements.txt        # Python dependencies
```

---

## Building

`build/compile-public-chimera-toolbox.ps1` compiles `src/app/obtoolbox_tui.py`
with Nuitka and packages it into `chimera-toolbox.exe` / `chimera-toolbox.zip`
at the repo root - this is the exact artifact meant to be attached to a
GitHub Release here, since it's what `chimera-toolbox.ps1` downloads via the
`irm | iex` one-liner.

This build has no login step and bundles no credentials, by design: this
repo is public, so nothing requiring authentication belongs in it.

Requires [Nuitka](https://nuitka.net/) and a C compiler (MSVC/clang) on
PATH. For a smaller compiled exe, drop a [UPX](https://upx.github.io/)
release anywhere under `build/` (e.g. `build/upx-5.2.1-win64/upx.exe`) -
the script finds it automatically, falling back to PATH, and just skips
the compression step with a warning if UPX isn't found at all.
