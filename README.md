# OverBytes Toolbox

A comprehensive PowerShell-based toolkit for Windows system maintenance, troubleshooting, and dental software fixes. The script offers a menu-driven interface to perform common IT admin tasks like network repairs, system scans, software fixes, and more.

---

## Features

- Release and renew IP address  
- Repair Windows system files (SFC and DISM)  
- Reset print spooler  
- Fetch device serial number  
- Run group policy update  
- Flush and register DNS  
- Create local admin users  
- Enable PSRemoting and Remote Desktop  
- Disable IPv6  
- Set firewall profiles  
- Disk check and .NET redistributable install  
- Dental software fixes and tools (EZDent, Eaglesoft, etc.)  
- Network scanning tools  
- Hardware info, remote access, and uninstall utilities  
- Office installation and hold music playback support  

---

## Usage

Run the script directly from PowerShell with the following command:

```powershell

irm overbytestech.com/repair | iex
```
If you run into SSL errors, run the following command:

```powershell

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; irm overbytestech.com/repair | iex
```
