Clear-Host
# Importing the modules
Write-Host "Setting temporary execution policy"
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
Write-Host "Setting TLS settings for script execution"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$testing = $false
$modulearray = @(
    "https://github.com/treyob/lib/releases/download/v0.2/dentalsoftware.psm1",
    "https://github.com/treyob/lib/releases/download/v0.2/ob.psm1",
    ""
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
    Invoke-WebRequest -Uri $psm1Url -OutFile $psm1FilePath 
    Import-Module -Name $psm1FilePath
    Write-Host "Imported Module $fileName"
}
if (-Not $testing) {
    ForEach ($psm1Url in $modulearray) {
        Import-OBModules $psm1Url
    }
} else {
    Write-Warning "Testing: Enabled"
    pause
    Import-OBModules "https://github.com/obmmoreno/obwinrepair/releases/download/obwinrepair-v0.3/ob.psm1"
    Import-OBModules "https://github.com/obmmoreno/obwinrepair/releases/download/obwinrepair-v0.3/dentalsoftware.psm1"
}

$global:sound = $null

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
"@ -ForegroundColor DarkCyan

Write-Host "7. Create Local User and Add to Administrators" -ForegroundColor DarkCyan -NoNewLine;    Write-Host "    |\     _,,,---,,_       " -ForegroundColor Cyan -NoNewLine; Write-Host "25. Revo Uninstaller" -ForegroundColor DarkCyan
Write-Host "8. Enable PSRemoting                       "    -ForegroundColor DarkCyan -NoNewLine; Write-Host "ZZZzz /, ```.-'```'    -.  ;-;;,_  " -ForegroundColor Cyan -NoNewLine; Write-Host "26. Office Install" -ForegroundColor DarkCyan
Write-Host "9. Disable IPv6                            "    -ForegroundColor DarkCyan -NoNewLine; Write-Host "     |,4-  ) )-,_.  ,\ (   `'-`' " -ForegroundColor Cyan -NoNewLine; Write-Host "27. Hold Music Section" -ForegroundColor DarkCyan
Write-Host "10. Enable Remote Desktop                  "    -ForegroundColor DarkCyan -NoNewLine; Write-Host "    `'---`'`'(_/--`'  ```-`' \_) " -ForegroundColor Cyan

Write-Host @"
11. Set Firewall Profiles to Allow All
12. CheckDisk C:\
13. Download and Install Microsoft Visual C++ 2015-2022 
"@ -ForegroundColor DarkCyan
Write-Host "14. Dental Software Section                 " -ForegroundColor DarkCyan -NoNewLine; Write-Host '     ___' -ForegroundColor Cyan
Write-Host "15. Server Power and Bitlocker Options      " -ForegroundColor DarkCyan -NoNewLine; Write-Host ' _.-|   |          |\__/,|   (`\' -ForegroundColor Cyan
Write-Host "16. Set Adobe Security Settings             " -ForegroundColor DarkCyan -NoNewLine; Write-Host '{   |   |          |o o  |__ _) )' -ForegroundColor Cyan
Write-Host "17. Get Domain Computers Info from DC       " -ForegroundColor DarkCyan -NoNewLine; Write-Host ' "-.|___|        _.( T   )  `  /' -ForegroundColor Cyan
Write-Host "18. Support Numbers                         " -ForegroundColor DarkCyan -NoNewLine; Write-Host "   .--'-`-.     _((_ `^--' /_<  \\" -ForegroundColor Cyan
Write-Host "00. Exit                                    " -ForegroundColor DarkCyan -NoNewLine; Write-Host '.+|______|__.-||__)`-''(((/  (((/' -ForegroundColor Cyan

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
# Practice Management Systems
$practiceSystems = @(
    [PSCustomObject]@{ Name = 'Eaglesoft';              Phone = '800-475-5036' }
    [PSCustomObject]@{ Name = 'Dentrix';                Phone = '800-336-8749' }
    [PSCustomObject]@{ Name = 'Open Dental';            Phone = '503-363-5432' }
    [PSCustomObject]@{ Name = 'Softdent';               Phone = '866-435-7473' }
    [PSCustomObject]@{ Name = 'Practiceworks';          Phone = '800-603-4438' }
    [PSCustomObject]@{ Name = 'Curve';                  Phone = '888-910-4376' }
    [PSCustomObject]@{ Name = 'OrthoEdge2';             Phone = '800-346-4504' }
    [PSCustomObject]@{ Name = 'Carestack';              Phone = '407-833-6123' }
    [PSCustomObject]@{ Name = 'TDO';                    Phone = '858-558-3696' }
)

Write-Host "Practice Management Systems" -ForegroundColor Green
$practiceSystems | Format-Table -AutoSize

# Imaging Systems
$imagingSystems = @(
    [PSCustomObject]@{ Name = 'DTX';                           Phone = '833-389-2255' }
    [PSCustomObject]@{ Name = 'Sirona (Sidexis\Schick)';       Phone = '800-659-5977' }
    [PSCustomObject]@{ Name = 'Apteryx';                       Phone = '800-861-5098' }
    [PSCustomObject]@{ Name = 'Vatech EZDent';                 Phone = '888-396-6872' }
    [PSCustomObject]@{ Name = 'Invivo';                        Phone = '888-883-3947' }
    [PSCustomObject]@{ Name = 'SmartScan Studio';              Phone = '888-883-3947' }
    [PSCustomObject]@{ Name = 'Dexis 9 & 10';                  Phone = '888-883-3947' }
    [PSCustomObject]@{ Name = 'ImageXL';                       Phone = '866-450-6717' }
    [PSCustomObject]@{ Name = 'Vixwin';                        Phone = '888-883-3947' }
    [PSCustomObject]@{ Name = 'SOTA Imaging';                  Phone = '714-532-6100' }
    [PSCustomObject]@{ Name = 'JMorita\i-Dixel';               Phone = '800-831-3222' }
    [PSCustomObject]@{ Name = 'Sicat';                         Phone = '800-550-9961' }
    [PSCustomObject]@{ Name = 'Ray Medical (SmartDent)';       Phone = '800-976-4586' }
    [PSCustomObject]@{ Name = 'Acteon';                        Phone = '800-289-6367' }
)
Write-Host "`nImaging Software" -ForegroundColor Green
$imagingSystems | Format-Table -AutoSize
        Read-Host "Press Enter to go back"}
    "19" { Invoke-NetScan }
    "20" { Start-Process powershell.exe -ArgumentList "-NoProfile -WindowStyle Hidden -Command & {Import-Module '$env:TEMP\obsoftware\ob.psm1'; Get-IPScanner; pause}" -Verb RunAs
        Read-Host "Advanced IP Scanner is downloading in the background and will start in a moment. " }
    "21" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-WizTree; Exit } > $null
        Read-Host "WizTree is downloading in the background and will start in a moment"
        }
    "22" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-DotNetRepair; Exit } > $null
        Read-Host ".NET Repair tool is downloading in the background and will start in a moment"
        }
    "23" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-HWInfo; Exit } > $null
        Read-Host "HWiNFO is downloading in the background and will start in a moment"
        }
    "24" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-TeamViewerQS; Exit } > $null 
        Read-Host "TeamViewerQS is downloading in the background and will start in a moment"
        }
    "25" { Start-Job -ScriptBlock { Import-Module "$env:TEMP\obsoftware\ob.psm1"; Invoke-RevoUninstaller; Exit } > $null 
        Read-Host "Revo Uninstaller is downloading in the background and will start in a moment" }
    "26" { Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Import-Module '$env:TEMP\obsoftware\ob.psm1'; Invoke-OfficeInstall; Exit" }
    "27" { Invoke-HoldMusicSection }
    "00" { Remove-Item -Recurse -Force "$env:TEMP\obsoftware"
        Write-Host "Cleaned up module files"
        Stop-HoldMusic
        $global:holdMusicPlaying = $false
        $choice = "000" }
    }
} while ($choice -ne "000")