Clear-Host
# Importing the modules

$modulearray = @(
    "https://github.com/treyob/lib/releases/download/v0.2/dentalsoftware.psm1",
    "https://github.com/treyob/lib/releases/download/v0.2/ob.psm1"
)
Function Import-OBModules {
    param($psm1Url)
    # Extract the file name from the URL
    $fileName = [System.IO.Path]::GetFileName($psm1Url)
    # Define the temp path and file path
    $tempPath = "$env:temp\obsoftware"
    $psm1FilePath = Join-Path -Path $tempPath -ChildPath $fileName
    # Create the directory if it doesn't exist
    if (-not (Test-Path -Path $tempPath)) {
        New-Item -Path $tempPath -ItemType Directory -Force > $null
    }
    Invoke-WebRequest -Uri $psm1Url -OutFile $psm1FilePath > $null
    Import-Module -Name $psm1FilePath
    Write-Host "Imported Module $fileName"
}
ForEach ($psm1Url in $modulearray) {
    Import-OBModules $psm1Url
}

do {
    # Set the PowerShell window/tab title
    $host.ui.RawUI.WindowTitle = "OverBytes Toolbox"
    Clear-Host
    Write-Host @"
________                   __________          __                     .____    .____   _________  
\_____  \___  __ __________\______   \___.__._/  |_  ____   ______    |    |   |    |  \_   ___ \ 
 /   |   \  \/ // __ \_  __ \    |  _<   |  |\   __\/ __ \ /  ___/    |    |   |    |  /    \  \/ 
/    |    \   /\  ___/|  | \/    |   \\___  | |  | \  ___/ \___ \     |    |___|    |__\     \____
\_______  /\_/  \___  |__|  |______  // ____| |__|  \___  >____  > /\ |_______ \_______ \______  /
        \/          \/             \/ \/                \/     \/  \/         \/       \/      \/ 
"@ -ForegroundColor Gray

    Write-Host "Choose an option:" -ForegroundColor DarkCyan
    Write-Host @"
1. Release and Renew IP Address                                           19. Run CLI Net Scan        
2. Repair Windows (System File Checker and DISM)      /\_/\               20. Advanced IPScanner             
3. Reset Print Spooler                               ( o.o )              21. WizTree 
4. Get Device Serial Number                           > ^ <               22. .Net Repair Tool
5. Run GPUpdate                                                           23. HWiNFO
6. Flush and Register DNS                                                 24. TeamViewerQS
7. Create Local User and Add to Administrators          |\     _,,,---,,_   
8. Enable PSRemoting                             ZZZzz /,```.-'```'    -.  ;-;;,_
9. Disable IPv6                                       |,4-  ) )-,_. ,\ (  `'-'
10. Enable Remote Desktop                            '---''(_/--'  ```-' \_) 
11. Set Firewall Profiles to Allow All
12. CheckDisk C:\
13. Download and Install Microsoft Visual C++ 2015-2022
14. Dental Software Section
15. Server Power and Bitlocker Options
16. Set Adobe Security Settings
17. Get Domain Computers Info from DC
18. Support Numbers
00. Exit 
"@ -ForegroundColor DarkCyan
    $choice = Read-Host "Enter the number of your choice"

# Main Menu Switch
switch ($choice) {
    "1" { Invoke-ReleaseRenew }
    "2" { Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Import-Module '$env:TEMP\obsoftware\ob.psm1'; Invoke-WindowsRepair; Exit" }
    "3" { Invoke-ResetSpooler }
    "4" { Get-DeviceInfo }
    "5" { Update-GroupPolicy }
    "6" { Invoke-DNSFlushRegister }
    "7" { New-LocalAdmin }
    "8" { Invoke-PSRemoting }
    "9" { Disable-IPv6 }
    "10" { Enable-RDP }
    "11" { Unlock-FirewallAll }
    "12" { Invoke-ChkdskC }
    "13" { Get-CppRedist }
    "14" { Invoke-DentalSoftwareSection }
    "15" { Invoke-ServerPowerOptions }
    "16" { Set-AdobeSettings; Invoke-UltraCat "Press Enter to Continue"; Read-Host }
    "17" { Get-ADDeviceInfo; Pause }
    "18" { Clear-Host
        Write-Host "****Practice Management****" -ForegroundColor Green
        Write-Host @"
Eaglesoft - 800-475-5036
Dentrix - 800-336-8749
Open Dental - 503-363-5432
Softdent - 866-435-7473
PracticeWorks - 800-603-4438
Curve - 888-910-4376
OrthoEdge2 - 800-346-4504
CareStack - 407-833-6123
TDO - 858-558-3696 
"@ -ForegroundColor DarkCyan
        Write-Host "****Imaging****" -ForegroundColor Green
        Write-Host @"
DTX - 833-389-2255
Sirona (Schick\Sidexis) - 800-861-5098
Apteryx - 800-861-5098
Vatech (EZDent) - 888-396-6872
Invivio 888-883-3947
SmartScan - 888-883-3947
Dexis - 888-883-3947
ImageXL - 866-450-6717
Vixwin - 888-883-3497
SOTA Imaging - 714-523-6100
SiCat - 800-550-9961
Carestream - 866-724-6317
Ray Medical - 800-976-4586
"@ -ForegroundColor DarkCyan
        Read-Host "Press Enter to go back"}
    "19" { Invoke-NetScan }
    "20" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Get-IPScanner; Exit } > $null
        Read-Host "IPScanner is downloading in the background and will start in a moment"
        }
    "21" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-WizTree; Exit } > $null
        Read-Host "WizTree is downloading in the background and will start in a moment"
        }
    "22" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-DotNetRepair; Exit } > $null
    Read-Host "HWiNFO is downloading in the background and will start in a moment"
        }
    "23" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-HWInfo; Exit } > $null
        Read-Host "HWiNFO is downloading in the background and will start in a moment"
        }
    "24" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-TeamViewerQS; Exit } > $null 
        Read-Host "TeamViewerQS is downloading in the background and will start in a moment"
        }
    "00" { Remove-Item -Recurse -Force "$env:TEMP\obsoftware"; $choice = "000"}
    }
} while ($choice -ne "000")