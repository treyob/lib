function Invoke-OfficeInstall {
    function Show-MainMenu {
        Clear-Host
        Write-Host @"
1. Microsoft 365 Business Standard
2. Office 2024 Home & Business
3. Office 2021 Home & Business
4. Office 2019 Home & Business
5. Office 2016 Home & Business (Oct. 14 2025 EOL)
6. Other
"@ -ForegroundColor DarkCyan
        return Read-Host "Enter the office version of your choice (1-6)"
    }

    do {
        $versionChoice = Show-MainMenu
    } while ($versionChoice -notin 1..6)

    if ($versionChoice -eq 6) {
        $versionArray = @(
            "AccessRetail",
            "Access2019Retail",
            "Access2021Retail",
            "Access2024Retail",
            "Access2019Volume",
            "Access2021Volume",
            "Access2024Volume",
            "ExcelRetail",
            "Excel2019Retail",
            "Excel2021Retail",
            "Excel2024Retail",
            "Excel2019Volume",
            "Excel2021Volume",
            "Excel2024Volume",
            "HomeBusinessRetail",
            "HomeBusiness2019Retail",
            "HomeBusiness2021Retail",
            "HomeBusiness2024Retail",
            "HomeStudentRetail",
            "HomeStudent2019Retail",
            "HomeStudent2021Retail",
            "Home2024Retail",
            "O365HomePremRetail",
            "OneNoteFreeRetail",
            "OneNoteRetail",
            "OneNote2021Volume",
            "OneNote2024Volume",
            "OutlookRetail",
            "Outlook2019Retail",
            "Outlook2021Retail",
            "Outlook2024Retail",
            "Outlook2019Volume",
            "Outlook2021Volume",
            "Outlook2024Volume",
            "Personal2019Retail",
            "Personal2021Retail",
            "PowerPointRetail",
            "PowerPoint2019Retail",
            "PowerPoint2021Retail",
            "PowerPoint2024Retail",
            "PowerPoint2019Volume",
            "PowerPoint2021Volume",
            "PowerPoint2024Volume",
            "ProfessionalRetail",
            "Professional2019Retail",
            "Professional2021Retail",
            "Professional2024Retail",
            "ProjectProXVolume",
            "ProjectPro2019Retail",
            "ProjectPro2021Retail",
            "ProjectPro2024Retail",
            "ProjectPro2019Volume",
            "ProjectPro2021Volume",
            "ProjectPro2024Volume",
            "ProjectStdRetail",
            "ProjectStdXVolume",
            "ProjectStd2019Retail",
            "ProjectStd2021Retail",
            "ProjectStd2024Retail",
            "ProjectStd2019Volume",
            "ProjectStd2021Volume",
            "ProjectStd2024Volume",
            "ProPlus2019Volume",
            "ProPlus2021Volume",
            "ProPlus2024Volume",
            "ProPlusSPLA2021Volume",
            "ProPlus2019Retail",
            "ProPlus2021Retail",
            "ProPlus2024Retail",
            "PublisherRetail",
            "Publisher2019Retail",
            "Publisher2021Retail",
            "Publisher2019Volume",
            "Publisher2021Volume",
            "Standard2019Volume",
            "Standard2021Volume",
            "StandardSPLA2021Volume",
            "Standard2024Volume",
            "VisioProXVolume",
            "VisioPro2019Retail",
            "VisioPro2021Retail",
            "VisioPro2024Retail",
            "VisioPro2019Volume",
            "VisioPro2021Volume",
            "VisioPro2024Volume",
            "VisioStdRetail",
            "VisioStdXVolume",
            "VisioStd2019Retail",
            "VisioStd2021Retail",
            "VisioStd2024Retail",
            "VisioStd2019Volume",
            "VisioStd2021Volume",
            "VisioStd2024Volume",
            "WordRetail",
            "Word2019Retail",
            "Word2021Retail",
            "Word2024Retail",
            "Word2019Volume",
            "Word2024Volume",
            "Word2021Volume",
            "O365ProPlusEEANoTeamsRetail",
            "O365ProPlusRetail",
            "O365BusinessEEANoTeamsRetail",
            "O365BusinessRetail"
        )

        $versionsHT = @{}
        for ($i = 0; $i -lt $versionArray.Count; $i++) {
            $key = ($i + 1).ToString()
            $versionsHT[$key] = $versionArray[$i]
        }

        $validInputs = $versionsHT.Keys + "exit"
        do {
            Clear-Host
            $versionsHT.GetEnumerator() | Sort-Object { [int]$_.Key } | ForEach-Object {
                Write-Host "$($_.Key): $($_.Value)"
            }
            $selectedKey = Read-Host "Enter the number or type 'exit' to cancel"
            if ($selectedKey -notin $validInputs) {
                Read-Host "Invalid input. Press Enter to try again"
            }
        } while ($selectedKey -notin $validInputs)

        if ($selectedKey -eq "exit") { return "Exiting..." }
        $selectedVersion = $versionsHT[$selectedKey]
    }
    else {
        $selectedVersion = $versionChoice
    }

    Clear-Host
    Install-Office -version $selectedVersion
}


function Install-Office {
    param ([string]$version)

    $basePath = "C:\officeodt"
    $xmlPath = "$basePath\officeodt-$version.xml"
    $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_18623-20156.exe"

    if (-not (Test-Path $basePath)) {
        New-Item -Path $basePath -ItemType Directory | Out-Null
        Write-Host "Created directory: $basePath"
    }

    $downloadMap = @{
        "1" = "O365BusinessRetail"
        "2" = "HomeBusiness2024Retail"
        "3" = "HomeBusiness2021Retail"
        "4" = "HomeBusiness2019Retail"
        "5" = "HomeBusinessRetail"
    }

    if ($version -in $downloadMap.Keys) {
        $productId = $downloadMap[$version]
        Write-Host "Downloading and installing $productId"
        Start-BitsTransfer -Source "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=$productId&platform=x64&language=en-us&version=O16GA" -Destination "$basePath\office.exe"        
        & "$basePath\office.exe"
        return
    }

    # Prompt for architecture
    do {
        $arch = Read-Host "Architecture? (32 or 64)"
    } while ($arch -notin @("32", "64"))

    # Exclude OneDrive?
    $excludeOneDrive = ""
    $excludeResponse = Read-Host "Would you like to exclude OneDrive? (Y/n)"
    if ($excludeResponse -notin @("n", "N", "no")) {
        $excludeOneDrive = '<ExcludeApp ID="OneDrive" />'
    }

    $xmlContent = @"
<Configuration>
    <Add SourcePath="$basePath\" OfficeClientEdition="$arch" Channel="Current">
        <Product ID="$version">
            <Language ID="en-us" />
            $excludeOneDrive
        </Product>
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="DisplayLevel" Value="Full" />
</Configuration>
"@

    $xmlContent | Out-File -FilePath $xmlPath -Encoding UTF8

    Start-BitsTransfer -Source $odtUrl -Destination "$basePath\officeodt.exe"
    Write-Host "Please accept the Microsoft License Terms and choose location" -NoNewline
    Write-Host " $basePath" -ForegroundColor Yellow
    & "$basePath\officeodt.exe"

    # Wait for setup
    do {
        $process = Get-Process -Name "officeodt" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)

    if (-not (Test-Path "$basePath\setup.exe")) {
        Write-Error "setup.exe not found in $basePath"
        Invoke-OfficeCleanUp
        return "Mission failed"
    }

    Write-Host "Downloading Office installer files..."
    & "$basePath\setup.exe" /download $xmlPath

    Write-Host "Installing Office..."
    & "$basePath\setup.exe" /configure $xmlPath

    Write-Host "Installation complete."
    Invoke-OfficeCleanUp
    Write-Host "Mission complete"
}

function Invoke-OfficeCleanUp {
    $path = "C:\officeodt"
    Write-Host "Cleaning up..."
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $path
    while (Test-Path $path) {
        Write-Host "Retrying removal of $path..."
        Start-Sleep -Seconds 3
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue $path
    }
}
Function Invoke-ReleaseRenew {
    Write-Host "Releasing and Renewing IP Address..." -ForegroundColor Magenta
    ipconfig /release
    ipconfig /renew
    Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Windows repair. Runs sfc and repair-windowsimage (DISM)
Function Invoke-WindowsRepair {
    Write-Host "Repairing Windows..." -ForegroundColor Magenta
    Write-Host "Running DISM to repair Windows image..." -ForegroundColor Magenta
    Repair-WindowsImage -Online -RestoreHealth
    Write-Host "Running System File Checker (sfc /scannow)..." -ForegroundColor Magenta
    sfc /scannow
    Write-Host "Repair completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Reset the printer spooler
Function Invoke-ResetSpooler {
    Write-Host "Resetting Print Spooler..." -ForegroundColor Magenta
    Stop-Service -Name spooler -Force
    Start-Service -Name spooler
    Write-Host "Print Spooler has been reset. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Get Device Serial Number
Function Get-DeviceInfo {
    Write-Host "Getting Device Info..." -ForegroundColor Magenta
    Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber
    Write-Host "Operation completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Group Policy update
Function Update-GroupPolicy {
    Write-Host "Running GPUpdate..." -ForegroundColor Magenta
    gpupdate /force
    Write-Host "Group Policy Update completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Flush and Register DNS
Function Invoke-DNSFlushRegister {
    Write-Host "Flushing and Registering DNS..." -ForegroundColor Magenta
    ipconfig /flushdns
    ipconfig /registerdns
    Write-Host "DNS operations completed. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Create Local Admin User
Function New-LocalAdmin {
    Write-Host "Creating a Local User and Adding to Administrators..." -ForegroundColor Magenta
    $username = Read-Host "Enter the username for the new user"
    $password = Read-Host "Enter the password for the new user" -AsSecureString
    New-LocalUser -Name $username -Password $password -FullName $username -Description "Local user created via script"
    Add-LocalGroupMember -Group "Administrators" -Member $username
    Write-Host "User $username has been created and added to the Administrators group. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Enable PS Remoting
Function Invoke-PSRemoting {
    Write-Host "Enabling PSRemoting..." -ForegroundColor Magenta
    Enable-PSRemoting -Force
    Write-Host "PSRemoting has been enabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Disable IPv6
Function Disable-IPv6 {
    Write-Host "Disabling IPv6..." -ForegroundColor Magenta
    Get-NetAdapter | ForEach-Object {
        Write-Host "Disabling IPv6 on adapter: $($_.Name)" -ForegroundColor Magenta
        Disable-NetAdapterBinding -Name $_.Name -ComponentID "ms_tcpip6"
    }
    Write-Host "IPv6 Disabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Enable Remote Desktop
Function Enable-RDP {
    Write-Host "Enabling Remote Desktop..." -ForegroundColor Magenta
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 0
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
    Write-Host "Remote Desktop has been enabled. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Allow all Windows Defender Firewall
Function Unlock-FirewallAll {
    Write-Host "Setting Firewall Profiles to Allow" -ForegroundColor Magenta
    Set-NetFirewallProfile -Profile Domain, Private, Public -DefaultInboundAction Allow
    Write-Host "Firewall configuration completed successfully. Press Enter to continue." -ForegroundColor Magenta
    Read-Host
}
# Run chkdsk on C
Function Invoke-chkdskC {
    Write-Host "Running CheckDisk for C:. A reboot will be required, can take up to 1 hour." -ForegroundColor Magenta
    chkdsk C: /r  
    Write-Host "Restart PC when it's best. Press Enter to continue." -ForegroundColor Magenta
    Read-Host     
}
# Download C++ Redistributables
Function Get-CppRedist {
    # Function to download and execute the redistributable installer
    function Install-Redistributable {
        param ($FileName, $Url)
        $filePath = "C:\$FileName"
        Invoke-WebRequest $Url -OutFile $filePath
        Write-Host "Running $FileName installer..." -ForegroundColor Magenta
        Start-Process $filePath
    }
    # Install both x64 and x86 C++ redistributables
    Install-Redistributable "c++.exe" "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    Install-Redistributable "c++x86.exe" "https://aka.ms/vs/17/release/vc_redist.x86.exe"
    Write-Host "Press Enter to go back (this will delete the app, make sure it's closed)" -ForegroundColor Magenta
    Read-Host
    Remove-Item C:\c++* -Force
}
# Run Advanced IP Scanner
Function Get-IPScanner {
    $zipUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/ipscanner.zip"
    $targetDir = "C:\ipscanner"
    Clear-Host; Write-Host "Downloading Advanced IP Scanner"
    Start-BitsTransfer -Source $zipUrl -Destination "$env:TEMP\obsoftware\ipscanner.zip"
    Clear-Host; Write-Host "Extracting Advanced IP Scanner"
    Expand-Archive -Path "$env:TEMP\obsoftware\ipscanner.zip" -DestinationPath $targetDir -Force
    Clear-Host; Write-Host "Running ipscanner"
    Start-Process -FilePath "$targetDir\Advanced_IP_Scanner_2.5.4594.1.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "Advanced IP scanner should now be running." "This screen will exit when you close it."
    do {
        $process = Get-Process -Name "Advanced_IP_Scanner_2.5.4594.1" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\obsoftware\ipscanner.zip"
}
# This function can determine the state of Bitlocker
Function Get-BitlockerStatus {
    # Get Protection Status
    $protectionOn = manage-bde -status C: | Select-String "Protection Status"
    
    if ($protectionOn -match "Protection Status:\s*Protection On") {
        $protectionOn = $true
    }
    elseif ($protectionOn -match "Protection Status:\s*Protection Off \(1 reboots left\)") {
        $protectionOn = "suspendreboot"
    }
    elseif ($protectionOn -match "Protection Status:\s*Protection Off") {
        $protectionOn = $false
    }

    # Get Encryption Status
    $fullyEncrypted = manage-bde -status C: | Select-String "Percent Encrypted"
    if ($fullyEncrypted -match "Percent Encrypted:\s* 100%\s*") {
        $fullyEncrypted = $true
    }
    else {
        $fullyEncrypted = $false
    }

    # Get Encryption Method
    $encryptionEnabled = manage-bde -status C: | Select-String "Encryption Method"
    if ($encryptionEnabled -notmatch "\s*Encryption Method:\s*None\s*") {
        $encryptionEnabled = $true
    }
    else {
        $encryptionEnabled = $false
    }

    # Determine BitLocker Status
    if ($protectionOn -eq $true -and $encryptionEnabled) {
        $bitlockerStatus = "protected"
    }
    elseif ($protectionOn -eq $false -and $encryptionEnabled) {
        $bitlockerStatus = "suspended"
    }
    elseif ($protectionOn -eq "suspendreboot") {
        $bitlockerStatus = "suspendreboot"
    }
    elseif ($fullyEncrypted -eq $false -and $protectionOn -eq $false -and $encryptionEnabled -eq $false) {
        $bitlockerStatus = "disabled"
    }
    else {
        $bitlockerStatus = "unexpected"
        Write-Host "BitLocker is in an unexpected state. See below"
        manage-bde -status C:
    }
    
    return $bitlockerStatus
}
# This function is used for the server power options with power prevention if drive is protected
Function Set-BitLockerState {
    param (
        [string]$action
    )
    if ($bitlockerStatus -in @("suspended", "suspendreboot", "disabled") -and $action -in @("Shutdown", "Reboot")) {
        Write-Host "BitLocker is suspended, proceeding with $action." -ForegroundColor Green
        $cmd = if ($action -eq "Shutdown") { shutdown /s /t 0 /c '$action by OverBytesTech script' /f /d p:4:1 } else { shutdown /r /t 0 /c "$action by OverBytesTech script" /f /d p:4:1 }
        $cmd
        Write-Host "Server will now $action." -ForegroundColor Yellow
        exit
    }
    elseif ($action -eq "Suspend") {
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspendreboot") { Write-Host "Resuming first"; Resume-BitLocker -MountPoint "C:" }
        Write-Host "Suspending BitLocker for C: until manually enabled."
        Suspend-BitLocker -MountPoint "C:" -RebootCount 0
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspended") {
            Write-Host "BitLocker successfully suspended for C:." -ForegroundColor Green
        }
        else {
            Write-Host "Bitlocker Suspend Failed!" -ForegroundColor Red
        }
        Read-Host "Press Enter to dismiss"
    }
    elseif ($action -eq "SuspendReboot") {
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspended") { Write-Host "Resuming first"; Resume-BitLocker -MountPoint "C:" }
        Write-Host "Suspending BitLocker for C: for 1 reboot."
        Suspend-BitLocker -MountPoint "C:" -RebootCount 1
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "suspendreboot") {
            Write-Host "BitLocker successfully suspended for C:." -ForegroundColor Green
        }
        else {
            Write-Host "Bitlocker Suspend Failed!`nBitlockerStatus:" -ForegroundColor Red; Get-BitlockerStatus
        }
        Read-Host "Press Enter to dismiss"
    }
    elseif ($action -eq "Resume") {
        Write-Host "Resuming BitLocker for C:."
        Resume-BitLocker -MountPoint "C:"
        $bitlockerStatus = Get-BitlockerStatus
        if ($bitlockerStatus -eq "protected") {
            Write-Host "BitLocker successfully resumed for C:." -ForegroundColor Green
        }
        else {
            Write-Host "Failed to resume BitLocker for C:.`nBitlockerStatus: $bitlockerStatus" -ForegroundColor Red
        }
        Read-Host "Press Enter to dismiss"
    }
    elseif ($action -eq "Status") {
        $bitlockerStatus = Get-BitlockerStatus
        Write-Host = "Status variable: $bitlockerStatus`n`n"
        manage-bde -status C:
        Read-Host "Press Enter to dismiss"
    }
    else {
        Write-Host "BitLocker is not suspended. Cannot proceed with $action." -ForegroundColor Red
        Read-Host "Press Enter to dismiss"
    }
}
# Setting Adobe Security Settings
Function Set-AdobeSettings {
    $acrobatPath32 = "Software\Adobe\Acrobat Reader"
    $acrobatPath64 = "Software\Adobe\Adobe Acrobat"

    function Get-UserFromSID {
        param ($sid)
        try {
            $securityIdentifier = New-Object System.Security.Principal.SecurityIdentifier($sid)
            $account = $securityIdentifier.Translate([System.Security.Principal.NTAccount])
        }
        catch {
            Write-Warning "Could not resolve SID: $sid"
        }
        return $account.Value
    }
    function Set-SecuritySettings {
        param ($path, $key, $arch)
        # Check if the path exists, if not, create the path
        if (-not (Test-Path $path)) {
            New-Item -Path $path -Force
            Write-Host "Created registry path: $path" -ForegroundColor Yellow
        }
        # Check if the key exists, if not, create the key
        if ($null -eq (Get-ItemProperty -Path $path -Name $key -ErrorAction SilentlyContinue)) {
            Set-ItemProperty -Path $path -Name $key -Value 0
            Write-Host "Created $key key for $username for Acrobat $version $arch" -ForegroundColor DarkGreen
        }
        else {
            #Write-Host "Working on path $path"
            if ($path -match "TrustManager") { $version = $path -replace '^.*\\([^\\]+)\\TrustManager$', '$1' }
            else { $version = $path -replace '^.*\\([^\\]+)\\Privileged$', '$1' }
            Write-Host "Disabling $key for $username for Acrobat $version $arch"
            Set-ItemProperty -Path $path -Name $key -Value 0
        }
    } 
    # Iterate over every user registry profile
    Clear-Host
    Get-ChildItem -Path "Registry::HKEY_USERS\" | Where-Object { $_.Name -match '^HKEY_USERS\\S-1-5-21-' -and $_.Name -notmatch '_Classes$' } | ForEach-Object {
        $sid = $_.Name -replace '^HKEY_USERS\\', ''
        $username = Get-UserFromSID $sid
        Write-Host "Found user: $username" -ForegroundColor DarkGreen
        # Ensure the correct path to Acrobat 32-Bit and exclude Volatile Environment
        if (Test-Path -Path "Registry::HKEY_USERS\$sid\$acrobatPath32") {
            # Ensure the registry path is not part of Volatile Environment before proceeding
            if ($sid -notmatch '\\Volatile Environment$') {
                Get-ChildItem -Path "Registry::HKEY_USERS\$sid\$acrobatPath32" | ForEach-Object {
                    # Set security settings for each matching user profile
                    Set-SecuritySettings "Registry::$_\TrustManager" "32-Bit"
                    Set-SecuritySettings "Registry::$_\TrustManager" "32-Bit"
                    Set-SecuritySettings "Registry::$_\Privileged" "32-Bit"
                    Write-Host
                }
            }
            else {
                Write-Host "Volatile Environment detected for SID $sid, skipping this profile."
            }
        }
        else {
            Write-Host "Acrobat 32-Bit not found"
        }

        # Check for Acrobat 64-Bit and exclude Volatile Environment
        if (Test-Path -Path "Registry::HKEY_USERS\$sid\$acrobatPath64") {
            # Ensure the registry path is not part of Volatile Environment before proceeding
            if ($sid -notmatch '\\Volatile Environment$') {
                Get-ChildItem -Path "Registry::HKEY_USERS\$sid\$acrobatPath64" | ForEach-Object {
                    # Set security settings for each matching user profile
                    Write-Host
                    Set-SecuritySettings "Registry::$_\TrustManager" "bEnhancedSecurityStandalone" "64-bit"
                    Set-SecuritySettings "Registry::$_\TrustManager" "bEnhancedSecurityInBrowser" "64-Bit"
                    Set-SecuritySettings "Registry::$_\Privileged" "bProtectedMode" "64-Bit"
                }
            }
            else {
                Write-Host "Volatile Environment detected for SID $sid, skipping this profile."
            }
        }
        else {
            Write-Host "Acrobat 64-Bit not found"
        }

    }
}
# Verbose Cats
Function Invoke-PounceCat {
    param ($msg, $msg2)
    Write-Host @"
   ____
  (.   \
    \  |  
     \ |___(\--/)
   __/    (  . . )   -- $msg
  "'._.    '-.O.'       $msg2
       '-.  \ "|\
          '.,,/'.,,
"@ -ForegroundColor Magenta
}
Function Invoke-UltraCat {
    param ($msg, $msg2)
    Write-Host @"

         /\_/\  
        ( o.o )  -- $msg
         > ^ <      $msg2 
       /       \
      /  |   |  \
     (   |   |   )
      \  |   |  /
       \_|___|_/
"@ -ForegroundColor Magenta
}

Function Get-ADDeviceInfo {
    # Get all computers from AD
    $computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name
    # Create an array to store results
    $results = @()
    # Loop through each computer and gather system information
    foreach ($computer in $computers) {
        try {
            # Check if the computer is online
            if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
                # Get System Information
                $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computer
                $cpu = Get-CimInstance -ClassName Win32_Processor -ComputerName $computer
                $ram = [math]::Round(($os.TotalVisibleMemorySize / 1MB), 2)  # Convert KB to GB
                $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'" -ComputerName $computer
                $bios = Get-CimInstance -ClassName Win32_BIOS -ComputerName $computer

                # Store the results
                $results += [PSCustomObject]@{
                    ComputerName = $computer
                    OSVersion    = "$($os.Caption) ($($os.Version))"
                    CPU          = $cpu.Name
                    RAM_GB       = $ram
                    DiskSize_GB  = [math]::Round(($disk.Size / 1GB), 2)
                    SerialNumber = $bios.SerialNumber
                }
            }
            else {
                Write-Warning "$computer is offline or unreachable."
            }
        }
        catch {
            Write-Warning "Failed to get data from $computer. Error: $_"
        }
    }

    # Export results to CSV
    $results | Export-Csv -Path "C:\DomainComputersInfo.csv" -NoTypeInformation

    # Display results
    $results
}

# Server Power and Bitlocker Options
Function Invoke-ServerPowerOptions {
    do {
        Clear-Host
        Write-Host "Server Power Options" -ForegroundColor Magenta
        $powerChoice = Read-Host @"
1. Shutdown now 
2. Reboot now 
3. Suspend BitLocker for C: for 1 reboot 
4. Suspend BitLocker for C: until manually enabled
5. Resume BitLocker for C:
6. Check C: BitLocker Status
7. Exit 

Enter the number of your choice
"@
        $bitlockerStatus = Get-BitlockerStatus
        switch ($powerChoice) {
            "1" { Set-BitLockerState -action "Shutdown" }
            "2" { Set-BitLockerState -action "Reboot" }
            "3" { Set-BitLockerState -action "SuspendReboot" }
            "4" { Set-BitLockerState -action "Suspend" }
            "5" { Set-BitLockerState -action "Resume" }
            "6" { Set-BitLockerState -action "Status" }
            "7" { Write-Host "Exiting Server Power Options" }
            Default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                Read-Host "Press Enter to try again"
            }
        }
    } while ($powerChoice -ne "7")
    
}
# Download, extract, and run wiztree
Function Invoke-WizTree {
    $zipUrl = "https://www.diskanalyzer.com/files/wiztree_4_24_portable.zip"
    $targetDir = "C:\WizTree"
    Clear-Host; Write-Host "Downloading WizTree"
    #Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\wiztree.zip"
    Start-BitsTransfer -Source $zipUrl -Destination "$env:TEMP\wiztree.zip"
    Clear-Host; Write-Host "Extracting WizTree"
    Expand-Archive -Path "$env:TEMP\wiztree.zip" -DestinationPath $targetDir
    Clear-Host; Write-Host "Running WizTree."
    Start-Process -FilePath "$targetDir\WizTree64.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "WizTree should now be running." "This screen will exit when you close WizTree."
    do {
        $process = Get-Process -Name "WizTree64" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\wiztree.zip"
}
Function Invoke-HWInfo {
    $zipUrl = "https://cytranet-dal.dl.sourceforge.net/project/hwinfo/Windows_Portable/hwi_820.zip?viasf=1"
    $targetDir = "C:\HWInfo"
    
    Clear-Host; Write-Host "Downloading hwinfo"
    #Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\hwinfo.zip"
    Start-BitsTransfer -Source $zipUrl -Destination "$env:TEMP\hwinfo.zip"
    Clear-Host; Write-Host "Extracting hwinfo"
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    Expand-Archive -Path "$env:TEMP\hwinfo.zip" -DestinationPath $targetDir -Force

    # Create the INI file directly
    $iniContent = @"
[Settings]
SummaryOnly=1
Theme=1
AutoUpdateBetaDisable=1
AutoUpdate=0
"@
    Set-Content -Path "$targetDir\HWiNFO64.INI" -Value $iniContent -Encoding ASCII

    Clear-Host; Write-Host "Running hwinfo."
    Start-Process -FilePath "$targetDir\hwinfo64.exe" -Verb RunAs
    
    Clear-Host; Invoke-PounceCat "hwinfo should now be running." "This screen will exit when you close hwinfo."
    
    do {
        $process = Get-Process -Name "hwinfo64" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\hwinfo.zip"
}

# Dot Net Repair Tool
Function Invoke-DotNetRepair {
    $installUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/NetFxRepairTool.exe"
    $targetDir = "C:\DotNetRepair"
    Clear-Host; Write-Host "Installing and Starting .Net Repair Tool"
    if (-Not (Test-Path -Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    Invoke-WebRequest -Uri $installUrl -OutFile "$targetDir\NetFxRepairTool.exe"
    Clear-Host; Write-Host "Running .Net Repair Tool."
    Start-Process -FilePath "$targetDir\NetFxRepairTool.exe" -Verb RunAs
    Clear-Host; Invoke-UltraCat ".Net Repair Tool should now be running." "This screen will exit when you close the tool."
    do {
        $process = Get-Process -Name "NetFxRepairTool" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}
Function Invoke-TDOXDRW11Fix {
    $installUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/TDOXDRWin11Fix.exe"
    $targetDir = "C:\TDOXDRWin11Fix"
    Clear-Host; Write-Host "Installing and Starting .Net Repair Tool"
    if (-Not (Test-Path -Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\TDOXDRWin11Fix.exe"
    Clear-Host; Write-Host "Downloading the script"
    Start-Process -FilePath "$targetDir\TDOXDRWin11Fix.exe" -Verb RunAs
    Clear-Host; Invoke-UltraCat "The fix should now be running." "This screen will exit when you close the installer."
    Start-Sleep -Seconds 2
    do {
        $process = Get-Process -Name "TDOXDRWin11Fix" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}
Function Invoke-NetScan {
    Clear-Host
    $Subnet = Read-Host "Enter subnet (e.g., 192.168.1)"
    $FirstIP = Read-Host "Enter the first IP in the range (e.g., 1)"
    $LastIP = Read-Host "Enter the first IP in the range (e.g., 254)"
    $OutputFile = "C:\ipscan.txt"
    $Results = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, [System.Environment]::ProcessorCount * 2)
    $RunspacePool.Open()
    $Runspaces = $FirstIP..$LastIP | ForEach-Object {
        $IP = "$Subnet.$_"
        $Runspace = [powershell]::Create().AddScript({
                param ($IP, $Results)
            
                # Test connectivity and retrieve details if online
                if (Test-Connection -ComputerName $IP -Count 1 -Quiet) {
                    $DeviceName = (Resolve-DnsName -Name $IP -ErrorAction SilentlyContinue).NameHost
                    $MAC = (Get-NetNeighbor -InterfaceAlias "*" | Where-Object { $_.IPAddress -eq $IP }).LinkLayerAddress

                    # Store online device details
                    $Results.Add([PSCustomObject]@{
                            "IP Address"  = $IP
                            "Device Name" = $DeviceName
                            "MAC Address" = $MAC
                        }) | Out-Null
                }
            }).AddArgument($IP).AddArgument($Results)

        $Runspace.RunspacePool = $RunspacePool
        [PSCustomObject]@{ Pipe = $Runspace; Status = $Runspace.BeginInvoke() }
    }
    $Runspaces | ForEach-Object {
        $_.Pipe.EndInvoke($_.Status)
        $_.Pipe.Dispose()
    }
    $RunspacePool.Close()
    $RunspacePool.Dispose()
    $SortedResults = $Results | Sort-Object { [int]($_."IP Address" -split "\.")[3] }
    $SortedResults | Format-Table -AutoSize
    $SaveToFile = Read-Host "Save results to file? (Y/N)"
    if ($SaveToFile -match "^[Yy]$") {
        $SortedResults | Format-Table -AutoSize | Out-File -Encoding UTF8 -FilePath $OutputFile
        Write-Host "Results saved to $OutputFile" -ForegroundColor Green
    }
    else {
        Write-Host "Results NOT saved." -ForegroundColor Yellow
    }
    Invoke-PounceCat "Press enter to continue"; Read-Host
}
Function Invoke-TeamViewerQS {
    $installUrl = "https://obtoolbox-public.s3.us-east-2.amazonaws.com/3rd-party-tools/TeamViewerQS_x64.exe"
    $targetDir = "$env:temp\obsoftware\TeamViewerQS"
    Clear-Host; Write-Host "Downloading and Starting TeamViewer"
    if (-Not (Test-Path -Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null}
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\TeamviewerQS.exe"
    Clear-Host; Write-Host "Running TeamViewer"
    Start-Process -FilePath "$targetDir\TeamViewerQS.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "TeamViewer should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "TeamViewer*" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-NewTeamViewerQS {
    $installUrl = "https://www.teamviewer.com/link/?url=505374"
    $targetDir = "C:\TeamViewerQS"
    Clear-Host; Write-Host "Downloading and Starting TeamViewer"
    if (-Not (Test-Path -Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\TeamViewerQS.exe"
    Clear-Host; Write-Host "Running TeamViewer"
    Start-Process -FilePath "$targetDir\TeamViewerQS.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "TeamViewer should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "TeamViewer*" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-DentrixInstallMigrateTool {
    $installUrl = "https://dentrix.com/support/core/MigrationAndInstallTool.exe"
    $targetDir = "C:\DentrixInstallMigrateTool"
    Clear-Host; Write-Host "Downloading and Starting Dentrix Installation and Migration tool"
    if (-Not (Test-Path -Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
    Start-BitsTransfer -Source $installUrl -Destination "$targetDir\DentrixInstallMigrateTool.exe"
    Clear-Host; Write-Host "Running the tool"
    Start-Process -FilePath "$targetDir\DentrixInstallMigrateTool.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "Dentrix Installation and Migration tool should now be running."
    do {
        $process = Get-Process | Get-Process | Where-Object { $_.Name -like "DentrixInstallMigrateTool" }
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir
}

Function Invoke-RevoUninstaller {
    Clear-Host
    $zipUrl = "https://81081bb5bd2a290e19d0-73f958688cc73e14784b0be099708265.ssl.cf1.rackcdn.com/RevoUninstaller_Portable.zip"
    $targetDir = "C:\RevoUninstaller"
    Clear-Host; Write-Host "Downloading RevoUninstaller"
    Invoke-WebRequest -Uri $zipUrl -OutFile "$env:TEMP\obsoftware\RevoUninstaller.zip"
    Clear-Host; Write-Host "Extracting RevoUninstaller"
    New-Item -ItemType Directory $targetDir -Force
    Expand-Archive -Path "$env:TEMP\obsoftware\RevoUninstaller.zip" -DestinationPath $targetDir -Force
    Clear-Host; Write-Host "Running RevoUninstaller."
    Start-Process -FilePath "$targetDir\RevoUninstaller_Portable\RevoUPort.exe" -Verb RunAs
    Clear-Host; Invoke-PounceCat "RevoUninstaller should now be running." "This screen will exit when you close RevoUninstaller."
    Start-Sleep -Seconds 3
    do {
        $process = Get-Process -Name "Revo*" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    } while ($process)
    Remove-Item -Recurse -Force $targetDir, "$env:TEMP\obsoftware\RevoUninstaller.zip"
}

Function Start-HoldMusic {
    param (
        [string]$link,
        [string]$songName
    )
    Stop-HoldMusic
    if (-Not (Test-Path -Path "$env:temp\obsoftware\$songName.wav")) {
        Write-Host "Downloading hold music"
        Start-BitsTransfer -Source "$link" -Destination "$env:temp\obsoftware\$songName.wav" > $null
    }
    $global:sound = New-Object System.Media.SoundPlayer "$env:temp\obsoftware\$songName.wav"
    $global:sound.PlayLooping()
}

Function Stop-HoldMusic {
    if ($null -ne $global:sound) {
        $global:sound.Stop()
        $global:sound.Dispose()
        $global:sound = $null
    }
}
function Invoke-HoldMusicSection {
    param ()
    do {
        Clear-Host
        Write-Host "Hold Music Section`n" -ForegroundColor DarkCyan
        $holdMusicTable = @(
            [PSCustomObject]@{ Option = '1. Other'; Song = 'Opus No. 1 - Tim Carleton' }
            [PSCustomObject]@{ Option = '2. Dexis'; Song = '12 Spanish Dances in G Major Arr. for Guitar' }
            [PSCustomObject]@{ Option = '3. Comcast'; Song = 'Winelight - Grover Washington' }
            [PSCustomObject]@{ Option = '4. Stop Music' }
        )
        $holdMusicTable | Format-Table -AutoSize

        switch ($musicChoice) {
            "1" { Start-HoldMusic -link 'https://obtoolbox-public.s3.us-east-2.amazonaws.com/hold-music/opus.no.1.wav' -songName 'Opus No. 1 - Tim Carleton' }
            "2" { Start-HoldMusic -link 'https://obtoolbox-public.s3.us-east-2.amazonaws.com/hold-music/12.span.dance.wav' -songName '12 Spanish Dances in G Major Arr. for Guitar' }
            "3" { Start-HoldMusic -link 'https://obtoolbox-public.s3.us-east-2.amazonaws.com/hold-music/Winelight+(152kbit_Opus).wav' -songName 'Winelight - Grover Washington' }
            "4" { Stop-HoldMusic }
        }
        $musicChoice = Read-Host "Enter the number of your choice, or press enter to exit"
    } while ($musicChoice -ne "")
}

Function Open-ExternalLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url
    )
    # Open with default browser associated with URLs using multiple fallbacks
    try {
        [System.Diagnostics.Process]::Start($Url) | Out-Null
    }
    catch {
        try {
            Start-Process -FilePath $Url -ErrorAction Stop
        }
        catch {
            Start-Process "cmd.exe" -ArgumentList "/c start `""" + $Url + "`""" -WindowStyle Hidden
        }
    }
}


### Start Sidexis Migration Functions
Function Set-SidexisServerPath {
    clear-host
    # Prompt user for new network share path
    $newPath = Read-Host "Enter the new Deployment Share path (e.g. \\SERVER02\PDATA)"

    # Validate format (basic check for UNC path)
    if ($newPath -notmatch "^\\\\[^\\]+\\[^\\]+") {
        Write-Host "Invalid path format. Please use a UNC path like \\SERVER02\PDATA." -ForegroundColor Red
        return
    }

    # Test network path accessibility
    Write-Host "Checking if the path is reachable..."
    if (!(Test-Path $newPath)) {
        Write-Host "The path '$newPath' is not reachable. Please check the network or path." -ForegroundColor Red
        return
    }

    # Set registry path
    $regPath = "HKLM:\SOFTWARE\Sirona\SIDEXIS4\Provisioning"

    # Check if key exists
    if (!(Test-Path $regPath)) {
        Write-Host "Registry path not found: $regPath" -ForegroundColor Red
        return
    }

    # Update registry key
    try {
        Set-ItemProperty -Path $regPath -Name "DeploymentShare" -Value $newPath
        Write-Host "DeploymentShare updated successfully to '$newPath'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to update registry: $_" -ForegroundColor Red
        return
    }
}
Function Install-SSMS {
    param()
    $ssmsUrl = "https://aka.ms/ssmsfullsetup"
    $installerPath = "$env:TEMP\SSMS-Setup.exe"
    Write-Host "Installing SQL Server Management Studio"
    Start-BitsTransfer -Source $ssmsUrl -Destination $installerPath
    Start-Process -FilePath $installerPath -ArgumentList "/install", "/quiet", "/norestart" -Wait
    Remove-Item $installerPath
    Write-Output "SSMS installation complete."
}
Function Invoke-CheckSidexisSQLInstances {
    param()
    $instanceName = "SIDEXIS_SQL"
    $instanceExists = Get-Service | Where-Object { $_.DisplayName -like "*$instanceName*" }

    if ($instanceExists) {
        Write-Warning "SQL Instance '$instanceName' already exists."
        $response = (Write-Host @"
Either press ENTER, uninstall SIDEXIS_SQL, and run script again, or type "SKIP" to skip this part of the script. 
Uninstall by going to:
Control Panel > Uninstall Programs > Double Click 'Microsoft SQL Server 2017' > In 'instance to remove', choose 'SIDEXIS_SQL'?

"@)
        if ($response -eq 'SKIP') {
            # Uninstall SQL instance is not automated yet
        }
        else {
            Write-Output "Aborting installation."
            exit
        }
    }
    else {
        Write-Host "Verified no other SidexisSQL instance exists."
    }
}
Function Install-SIDEXIS_SQL {
    $installDir = "C:\SQL2017"
    $bootstrapper = "$installDir\SQLServer2017-SSEI-Expr.exe"
    $mediaDir = "$installDir\Media"
    $iniFile = "$installDir\ConfigurationFile.ini"
    $extractPath = "$mediaDir\Extracted"
    $instanceName = "SIDEXIS_SQL"
    $serviceName = "MSSQL`$$instanceName"

    if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }

    if (-not (Test-Path $bootstrapper)) {
        Write-Host "Downloading bootstrapper..."
        Start-BitsTransfer -Source "https://download.microsoft.com/download/5/e/9/5e9b18cc-8fd5-467e-b5bf-bade39c51f73/SQLServer2017-SSEI-Expr.exe" -Destination $bootstrapper
    }

    if (-not (Test-Path "$mediaDir\SQLEXPR_x64_ENU.exe")) {
        Write-Host "Downloading full media..."
        New-Item -Path $mediaDir -ItemType Directory -Force | Out-Null
        Start-Process -FilePath $bootstrapper -ArgumentList "/Action=Download /Quiet /MediaPath=`"$mediaDir`" /MediaType=Core" -Wait
    }

    @"
[OPTIONS]
ACTION="Install"
ROLE="AllFeatures_WithDefaults"
ENU="True"
QUIETSIMPLE="True"
UpdateEnabled="False"
USEMICROSOFTUPDATE="False"
FEATURES=SQLENGINE,REPLICATION,SNAC_SDK
INSTANCENAME="$instanceName"
INSTANCEID="$instanceName"
INSTALLSHAREDDIR="C:\Program Files\Microsoft SQL Server"
INSTALLSHAREDWOWDIR="C:\Program Files (x86)\Microsoft SQL Server"
INSTANCEDIR="C:\Program Files\Microsoft SQL Server"

SQLSVCACCOUNT="NT Service\MSSQL`$$instanceName"
SQLSVCSTARTUPTYPE="Automatic"
SQLSVCINSTANTFILEINIT="False"

AGTSVCACCOUNT="NT AUTHORITY\NETWORK SERVICE"
AGTSVCSTARTUPTYPE="Disabled"

SQLTELSVCACCT="NT Service\SQLTELEMETRY`$$instanceName"
SQLTELSVCSTARTUPTYPE="Automatic"

SECURITYMODE="SQL"
SAPWD="2BeChanged!"

ADDCURRENTUSERASSQLADMIN="True"

TCPENABLED="1"
NPENABLED="1"
BROWSERSVCSTARTUPTYPE="Automatic"

SQLCOLLATION="SQL_Latin1_General_CP1_CI_AS"
ENABLERANU="True"

SQLTEMPDBFILECOUNT="1"
SQLTEMPDBFILESIZE="8"
SQLTEMPDBFILEGROWTH="64"
SQLTEMPDBLOGFILESIZE="8"
SQLTEMPDBLOGFILEGROWTH="64"

IACCEPTSQLSERVERLICENSETERMS="True"
"@ | Set-Content -Path $iniFile -Encoding ASCII

    New-Item -Path $extractPath -ItemType Directory -Force | Out-Null
    Write-Host "Extracting setup files silently..."
    & "$mediaDir\SQLEXPR_x64_ENU.exe" /Q /IACCEPTSQLSERVERLICENSETERMS /ENU /X:"$extractPath"

    $setupExe = Join-Path $extractPath "SETUP.EXE"
    $timeout = 30
    $elapsed = 0
    while (-not (Test-Path $setupExe) -and $elapsed -lt $timeout) {
        Start-Sleep -Seconds 1
        $elapsed++
    }

    if (Test-Path $setupExe) {
        Write-Host "Running silent install..."
        & $setupExe /ConfigurationFile="$iniFile"

        # Wait a bit for service registration
        Start-Sleep -Seconds 10

        # Try starting the service
        $sqlServiceName = "MSSQL`$$instanceName"
        Write-Host "Checking SQL Server service: $sqlServiceName"
        $service = Get-Service -Name $sqlServiceName -ErrorAction SilentlyContinue

        if ($service -and $service.Status -ne 'Running') {
            try {
                Start-Service -Name $sqlServiceName -ErrorAction Stop
                Start-Sleep -Seconds 5
            }
            catch {
                Write-Warning "SQL Service failed to start. Attempting rebuild..."

                # Run RebuildDatabase
                C:\SQL2017\Media\Extracted\setup.exe /QUIET /ACTION=REBUILDDATABASE /INSTANCENAME=SIDEXIS_SQL /SQLSYSADMINACCOUNTS="BUILTIN\Administrators" /SAPWD="2BeChanged!" /IACCEPTSQLSERVERLICENSETERMS

                # Try starting again
                Start-Sleep -Seconds 5
                try {
                    Start-Service -Name $sqlServiceName
                    Write-Host "Service started after rebuild."
                }
                catch {
                    Write-Error "Rebuild failed or service could not be started."
                }
            }
        }
        elseif ($service -and $service.Status -eq 'Running') {
            Write-Host "SQL Server instance '$instanceName' is running."
        }
        else {
            Write-Error "Service $sqlServiceName not found."
        }
    }
    else {
        Write-Error "SETUP.EXE not found after waiting $timeout seconds. Install aborted."
    }
    Remove-Item -Recurse "C:\SQL2017"
}
function Get-LatestDBBackups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$dbBackupDir
    )

    # Ensure the directory exists
    if (-Not (Test-Path $dbBackupDir)) {
        Write-Error "Directory '$dbBackupDir' does not exist."
        return
    }

    # Get all .bak files in the directory
    $bakFiles = Get-ChildItem -Path $dbBackupDir -Filter *.bak -File

    # Get latest PDATA_SQLEXPRESS.bak
    $latestPdataBak = $bakFiles |
    Where-Object { $_.Name -like '*_PDATA_SQLEXPRESS.bak' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

    # Get latest SIDEXIS.bak
    $latestSidexisBak = $bakFiles |
    Where-Object { $_.Name -like '*_SIDEXIS.bak' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

    # Output results
    if ($latestPdataBak) {
        Write-Output "Latest PDATA_SQLEXPRESS.bak: $($latestPdataBak.Name) - Last Modified: $($latestPdataBak.LastWriteTime)"
    }
    else {
        Write-Warning "No PDATA_SQLEXPRESS.bak files found."
    }

    if ($latestSidexisBak) {
        Write-Output "Latest SIDEXIS.bak: $($latestSidexisBak.Name) - Last Modified: $($latestSidexisBak.LastWriteTime)"
    }
    else {
        Write-Warning "No SIDEXIS.bak files found."
    }

    # Return as output variables
    return @{
        PDATA_SQLEXPRESS = $latestPdataBak
        SIDEXIS          = $latestSidexisBak
    }
}
function Find-SqlCmd {
    Write-Host "Searching for sqlcmd.exe..."
    $paths = @(
        "$env:ProgramFiles\Microsoft SQL Server",
        "$env:ProgramFiles(x86)\Microsoft SQL Server",
        "$env:ProgramW6432\Microsoft SQL Server"
    )

    foreach ($base in $paths) {
        try {
            $result = Get-ChildItem -Path $base -Recurse -Filter sqlcmd.exe -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($result) {
                Write-Host "Found sqlcmd at: $($result.FullName)"
                return $result.FullName
            }
        }
        catch {}
    }

    Write-Warning "sqlcmd.exe not found. Manual restore required."
    return $null
}
Function Restore-DBBackups {

    function Get-LogicalFileNames {
        param (
            [string]$sqlcmdPath,
            [string]$sqlInstance,
            [string]$sqlUser,
            [string]$sqlPassword,
            [string]$bakPath
        )

        $query = "RESTORE FILELISTONLY FROM DISK = N'$bakPath';"
        $output = & "$sqlcmdPath" -S $sqlInstance -U $sqlUser -P $sqlPassword -Q $query -W -s"," 2>$null

        if (-not $output -or $output.Count -lt 2) {
            throw "Failed to retrieve logical file names from $bakPath"
        }

        $lines = $output | Where-Object { $_ -and ($_ -notmatch "^-+") }
        $data = ($lines[1] -split ',')[0].Trim('"')
        $log = ($lines[2] -split ',')[0].Trim('"')
        return @{ Data = $data; Log = $log }
    }

    function Restore-Database {
        param (
            [string]$dbName,
            [string]$bakPath
        )

        Write-Host "`nRestoring $dbName from $bakPath..."

        $sqlcmdPath = Find-SqlCmd
        if (-not $sqlcmdPath) {
            Write-Error "ERROR: sqlcmd.exe not found. Please restore '$dbName' manually using SSMS or T-SQL."
            return
        }

        $sqlInstance = "localhost\SIDEXIS_SQL"
        $sqlUser = "sa"
        $sqlPassword = "2BeChanged!"

        try {
            $logical = Get-LogicalFileNames -sqlcmdPath $sqlcmdPath -sqlInstance $sqlInstance -sqlUser $sqlUser -sqlPassword $sqlPassword -bakPath $bakPath

            $mdfPath = "C:\Program Files\Microsoft SQL Server\MSSQL14.SIDEXIS_SQL\MSSQL\DATA\$dbName.mdf"
            $ldfPath = "C:\Program Files\Microsoft SQL Server\MSSQL14.SIDEXIS_SQL\MSSQL\DATA\${dbName}_log.ldf"

            $restoreQuery = @"
USE [master];
IF DB_ID('$dbName') IS NOT NULL
BEGIN
    ALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE [$dbName];
END
RESTORE DATABASE [$dbName]
FROM DISK = N'$bakPath'
WITH MOVE N'$($logical.Data)' TO N'$mdfPath',
    MOVE N'$($logical.Log)' TO N'$ldfPath',
    REPLACE;
"@

            $tempSqlFile = New-TemporaryFile
            Set-Content -Path $tempSqlFile.FullName -Value $restoreQuery -Encoding UTF8

            & "$sqlcmdPath" -S $sqlInstance -U $sqlUser -P $sqlPassword -i $tempSqlFile.FullName
            if ($LASTEXITCODE -ne 0) {
                throw "sqlcmd returned exit code $LASTEXITCODE"
            }

            Write-Host "$dbName restored successfully."
        }
        catch {
            Write-Error "❌ Failed to restore $dbName. Please restore it manually. Details: $_"
        }
        finally {
            if ($tempSqlFile) { Remove-Item $tempSqlFile.FullName -Force -ErrorAction SilentlyContinue }
        }
    }

    # === Begin Script ===

    $backupDir = "G:\PDATA\Backup"
    Write-Host "Backup directory set to: $backupDir"

    $PDATA_SQLEXPRESS = Get-ChildItem -Path $backupDir -Filter "*_PDATA_SQLEXPRESS.bak" | Select-Object -First 1
    $SIDEXIS = Get-ChildItem -Path $backupDir -Filter "*_SIDEXIS.bak" | Select-Object -First 1

    Write-Host "Path 1: $($SIDEXIS.FullName)"
    Write-Host "Path 2: $($PDATA_SQLEXPRESS.FullName)"

    if ($PDATA_SQLEXPRESS) {
        Restore-Database -dbName "PDATA" -bakPath $($PDATA_SQLEXPRESS.FullName)
    }
    else {
        Write-Warning "❌ Could not find a PDATA_SQLEXPRESS .bak file."
    }

    if ($SIDEXIS) {
        Restore-Database -dbName "SIDEXIS" -bakPath $($SIDEXIS.FullName)
    }
    else {
        Write-Warning "❌ Could not find a SIDEXIS .bak file."
    }
}
function Remove-ProvisioningJobTargetType0 {
    param(
        [string]$sqlInstance,
        [string]$sqlUser,
        [string]$sqlPassword,
        [string]$database = "SIDEXIS"
    )

    $sqlcmdPath = Find-SqlCmd
    if (-not $sqlcmdPath) {
        Write-Error "sqlcmd.exe not found on this machine. Please install SQL Server Command Line Utilities."
        return
    }

    # Create temp SQL file to delete the row with TargetType = 0
    $sql = @"
USE [$database];
DELETE FROM dbo.ProvisioningJob WHERE TargetType = 0;
GO
"@

    $tempSqlFile = [System.IO.Path]::GetTempFileName() + ".sql"
    Set-Content -Path $tempSqlFile -Value $sql -Encoding UTF8

    try {
        # Run sqlcmd with provided credentials
        $args = @(
            "-S", $sqlInstance,
            "-U", $sqlUser,
            "-P", $sqlPassword,
            "-i", $tempSqlFile
        )
        $process = Start-Process -FilePath $sqlcmdPath -ArgumentList $args -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Host "Row(s) with TargetType=0 deleted from dbo.ProvisioningJob successfully."
        }
        else {
            Write-Error "Failed to delete row from dbo.ProvisioningJob. Exit code: $($process.ExitCode)"
        }
    }
    catch {
        Write-Error "Error running sqlcmd: $_"
    }
    finally {
        Remove-Item -Path $tempSqlFile -ErrorAction SilentlyContinue
    }
}
Function Open-SidexisMigrationSection {
    do {
        Clear-Host
        Write-Warning "For production, do not use on Windows Server 2025. SQL 2017 is not officially compatible"
        Write-Host "Sidexis Database Migration`n" -ForegroundColor DarkCyan
        Write-Host "Scripts (Est. 20-45 min)" -ForegroundColor DarkYellow
        Write-Host @"
1. Sql Server Management Software Installer              _._     _,-'""```-._
2. SIDEXIS_SQL Installer                               (,-.```._,'(       |\```-/| 
3. Restore from DB Backup and configure tool                ```-.-' \ )-```( , o o)
4. Manually Install Sidexis Server                                 ```-    \```_`"'-"
5. Provision each workstation to use the new PDATA share (Run on each workstation)

"@
        Write-Host "For Aquisition Server and IO Software:" -ForegroundColor DarkYellow                                                
        Write-Host @" 
7. Sidexis Migration Documentation

"@                                               
        Write-Host @" 
00. Exit Sidexis Database Migration        
"@
        $dentalChoice = Read-Host "Enter the number of your choice"
    

        switch ($dentalChoice) {
            "1" { Start-Process powershell.exe -ArgumentList "-NoProfile -Command & {Import-Module '$env:TEMP\obsoftware\ob.psm1'; Install-SSMS}" -Verb RunAs }
            "2" { Start-Process powershell.exe -ArgumentList "-NoProfile -Command & {Import-Module '$env:TEMP\obsoftware\ob.psm1'; Install-SIDEXIS_SQL}" -Verb RunAs }
            "3" { Start-Process powershell.exe -ArgumentList "-NoProfile -Command & {Import-Module '$env:TEMP\obsoftware\ob.psm1'; Restore-DBBackups; Remove-ProvisioningJobTargetType0 -sqlInstance 'localhost\SIDEXIS_SQL' -sqlUser 'sa' -sqlPassword '2BeChanged!'; pause}" -Verb RunAs }
            "5" { Set-SidexisServerPath; Read-Host "Press enter to dismiss" }
            "7" { Open-ExternalLink "https://www.dentsplysironasupport.com/content/dam/master/product-procedure-brand-categories/imaging/product-categories/software/imaging-software/sidexis-4/Sidexis%204%20Migration%20Guide%20Rev.2.pdf" }
        }
    } while ($dentalChoice -ne "00")
}

### End Sidexis Migration Functions